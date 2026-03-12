#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use chrono::Utc;
use serde::Serialize;
use serde_json::{Map, Value};
use std::collections::HashSet;
use std::fs;
use std::io::ErrorKind;
use std::path::{Path, PathBuf};
use tauri::AppHandle;
use tauri::Manager;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SaveResult {
    path: String,
    saved_at: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct AppInfo {
    data_path: String,
    name: String,
    version: String,
}

fn legacy_data_file_path(app: &AppHandle) -> Result<PathBuf, String> {
    let data_dir = app.path().app_data_dir().map_err(|err| err.to_string())?;
    fs::create_dir_all(&data_dir).map_err(|err| err.to_string())?;
    Ok(data_dir.join("oneplace-data.json"))
}

fn data_root_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let data_dir = app.path().app_data_dir().map_err(|err| err.to_string())?;
    let root = data_dir.join("workspace");
    fs::create_dir_all(&root).map_err(|err| err.to_string())?;
    Ok(root)
}

fn workspace_manifest_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(data_root_dir(app)?.join("workspace.json"))
}

fn notebooks_root_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let root = data_root_dir(app)?.join("notebooks");
    fs::create_dir_all(&root).map_err(|err| err.to_string())?;
    Ok(root)
}

fn slugify(value: &str) -> String {
    let mut slug = String::new();
    let mut last_was_dash = false;

    for ch in value.chars() {
        let normalized = ch.to_ascii_lowercase();
        if normalized.is_ascii_alphanumeric() {
            slug.push(normalized);
            last_was_dash = false;
        } else if !last_was_dash {
            slug.push('-');
            last_was_dash = true;
        }
    }

    slug.trim_matches('-').to_string()
}

fn as_object_mut(value: &mut Value) -> Result<&mut Map<String, Value>, String> {
    value
        .as_object_mut()
        .ok_or_else(|| "Expected JSON object.".to_string())
}

fn get_str<'a>(value: &'a Value, key: &str) -> Result<&'a str, String> {
    value
        .get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| format!("Missing string field: {key}"))
}

fn write_pretty_json(path: &Path, value: &Value) -> Result<(), String> {
    let raw = serde_json::to_vec_pretty(value).map_err(|err| err.to_string())?;
    fs::write(path, raw).map_err(|err| err.to_string())
}

fn read_json_file(path: &Path) -> Result<Value, String> {
    let raw = fs::read_to_string(path).map_err(|err| err.to_string())?;
    serde_json::from_str(&raw).map_err(|err| err.to_string())
}

fn load_notebook_from_dir(notebook_dir: &Path) -> Result<Value, String> {
    let notebook_file = notebook_dir.join("notebook.json");
    let mut notebook = read_json_file(&notebook_file)?;

    let groups = as_object_mut(&mut notebook)?
        .get_mut("sectionGroups")
        .and_then(Value::as_array_mut)
        .ok_or_else(|| "Notebook metadata is missing section groups.".to_string())?;

    for group in groups {
        let sections = as_object_mut(group)?
            .get_mut("sections")
            .and_then(Value::as_array_mut)
            .ok_or_else(|| "Notebook metadata is missing sections.".to_string())?;

        for section in sections.iter_mut() {
            let pages_file = section
                .get("pagesFile")
                .and_then(Value::as_str)
                .ok_or_else(|| "Section metadata is missing pagesFile.".to_string())?;
            let section_path = notebook_dir.join(pages_file);
            *section = read_json_file(&section_path)?;
        }
    }

    Ok(notebook)
}

fn section_file_name(group_id: &str, section_id: &str) -> String {
    format!("{group_id}__{section_id}.json")
}

fn write_notebook_to_dir(notebook: &Value, notebook_dir: &Path, clean_dir: bool) -> Result<(), String> {
    let notebook_id = get_str(notebook, "id")?;

    if clean_dir && notebook_dir.exists() {
        fs::remove_dir_all(notebook_dir).map_err(|err| err.to_string())?;
    }

    let sections_dir = notebook_dir.join("sections");
    fs::create_dir_all(&sections_dir).map_err(|err| err.to_string())?;

    let mut notebook_metadata = notebook.clone();
    let notebook_object = as_object_mut(&mut notebook_metadata)?;
    let section_groups = notebook_object
        .get_mut("sectionGroups")
        .and_then(Value::as_array_mut)
        .ok_or_else(|| format!("Notebook {notebook_id} is missing section groups."))?;

    for group in section_groups {
        let group_id = get_str(group, "id")?.to_string();
        let group_sections = as_object_mut(group)?
            .get_mut("sections")
            .and_then(Value::as_array_mut)
            .ok_or_else(|| format!("Section group {group_id} is missing sections."))?;

        for section in group_sections.iter_mut() {
            let full_section = section.clone();
            let section_id = get_str(&full_section, "id")?.to_string();
            let filename = section_file_name(&group_id, &section_id);
            let section_path = sections_dir.join(&filename);
            write_pretty_json(&section_path, &full_section)?;

            let section_object = as_object_mut(section)?;
            section_object.remove("pages");
            section_object.insert("pagesFile".to_string(), Value::String(format!("sections/{filename}")));
        }
    }

    let notebook_file = notebook_dir.join("notebook.json");
    write_pretty_json(&notebook_file, &notebook_metadata)?;
    Ok(())
}

fn save_workspace(raw_data: &str, app: &AppHandle) -> Result<PathBuf, String> {
    let mut app_state: Value = serde_json::from_str(raw_data).map_err(|err| err.to_string())?;
    let notebooks = app_state
        .get("notebooks")
        .and_then(Value::as_array)
        .cloned()
        .ok_or_else(|| "App state is missing notebooks.".to_string())?;

    let manifest_path = workspace_manifest_path(app)?;
    let notebooks_root = notebooks_root_dir(app)?;
    let mut notebook_refs = Vec::with_capacity(notebooks.len());
    let mut active_dirs = HashSet::with_capacity(notebooks.len());

    for notebook in notebooks {
        let notebook_id = get_str(&notebook, "id")?;
        let notebook_name = get_str(&notebook, "name")?;
        let slug = slugify(notebook_name);
        let folder_name = if slug.is_empty() {
            format!("notebook-{notebook_id}")
        } else {
            format!("{slug}-{notebook_id}")
        };
        let notebook_dir = notebooks_root.join(&folder_name);
        write_notebook_to_dir(&notebook, &notebook_dir, true)?;

        active_dirs.insert(folder_name);
        notebook_refs.push(Value::Object(Map::from_iter([
            ("id".to_string(), Value::String(notebook_id.to_string())),
            (
                "path".to_string(),
                Value::String(format!(
                    "notebooks/{}/notebook.json",
                    notebook_dir
                        .file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or_default()
                )),
            ),
        ])));
    }

    for entry in fs::read_dir(&notebooks_root).map_err(|err| err.to_string())? {
        let entry = entry.map_err(|err| err.to_string())?;
        if !entry.file_type().map_err(|err| err.to_string())?.is_dir() {
            continue;
        }
        let name = entry.file_name().to_string_lossy().to_string();
        if !active_dirs.contains(&name) {
            fs::remove_dir_all(entry.path()).map_err(|err| err.to_string())?;
        }
    }

    let manifest_object = as_object_mut(&mut app_state)?;
    manifest_object.insert("notebooks".to_string(), Value::Array(notebook_refs));
    write_pretty_json(&manifest_path, &app_state)?;

    Ok(manifest_path)
}

fn load_workspace(app: &AppHandle) -> Result<Option<String>, String> {
    let manifest_path = workspace_manifest_path(app)?;
    match fs::read_to_string(&manifest_path) {
        Ok(raw_manifest) => {
            let mut manifest: Value = serde_json::from_str(&raw_manifest).map_err(|err| err.to_string())?;
            let notebook_refs = manifest
                .get("notebooks")
                .and_then(Value::as_array)
                .cloned()
                .ok_or_else(|| "Workspace manifest is missing notebook references.".to_string())?;
            let data_root = data_root_dir(app)?;
            let mut notebooks = Vec::with_capacity(notebook_refs.len());

            for notebook_ref in notebook_refs {
                let notebook_path = get_str(&notebook_ref, "path")?;
                let absolute_notebook_path = data_root.join(notebook_path);
                let notebook_dir = absolute_notebook_path
                    .parent()
                    .ok_or_else(|| "Notebook path is invalid.".to_string())?;
                notebooks.push(load_notebook_from_dir(notebook_dir)?);
            }

            as_object_mut(&mut manifest)?.insert("notebooks".to_string(), Value::Array(notebooks));
            let raw = serde_json::to_string(&manifest).map_err(|err| err.to_string())?;
            Ok(Some(raw))
        }
        Err(err) if err.kind() == ErrorKind::NotFound => {
            let legacy_path = legacy_data_file_path(app)?;
            match fs::read_to_string(&legacy_path) {
                Ok(raw) => Ok(Some(raw)),
                Err(err) if err.kind() == ErrorKind::NotFound => Ok(None),
                Err(err) => Err(err.to_string()),
            }
        }
        Err(err) => Err(err.to_string()),
    }
}

#[tauri::command]
fn load_data(app: AppHandle) -> Result<Option<String>, String> {
    load_workspace(&app)
}

#[tauri::command]
fn save_data(app: AppHandle, raw_data: String) -> Result<SaveResult, String> {
    let manifest_path = save_workspace(&raw_data, &app)?;

    Ok(SaveResult {
        path: manifest_path.display().to_string(),
        saved_at: Utc::now().to_rfc3339(),
    })
}

#[tauri::command]
fn get_app_info(app: AppHandle) -> Result<AppInfo, String> {
    let data_path = data_root_dir(&app)?;
    let package_info = app.package_info();

    Ok(AppInfo {
        data_path: data_path.display().to_string(),
        name: package_info.name.clone(),
        version: package_info.version.to_string(),
    })
}

#[tauri::command]
fn open_notebook_dir(path: String) -> Result<String, String> {
    let notebook = load_notebook_from_dir(Path::new(&path))?;
    serde_json::to_string(&notebook).map_err(|err| err.to_string())
}

#[tauri::command]
fn export_notebook_dir(path: String, notebook: String) -> Result<SaveResult, String> {
    let notebook_value: Value = serde_json::from_str(&notebook).map_err(|err| err.to_string())?;
    let notebook_name = get_str(&notebook_value, "name")?;
    let folder_name = slugify(notebook_name);
    let target_dir = if folder_name.is_empty() {
        Path::new(&path).join("notebook-export")
    } else {
        Path::new(&path).join(folder_name)
    };

    write_notebook_to_dir(&notebook_value, &target_dir, true)?;

    Ok(SaveResult {
        path: target_dir.display().to_string(),
        saved_at: Utc::now().to_rfc3339(),
    })
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            load_data,
            save_data,
            get_app_info,
            open_notebook_dir,
            export_notebook_dir
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
