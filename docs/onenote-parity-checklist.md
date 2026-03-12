# OneNote Parity Checklist

This project should be built toward deliberate OneNote parity in layers, not as isolated UI tweaks.

## Product Decision

- Target mode: exact OneNote desktop behavior where practical.
- Constraint: phase the work so navigation and data-model parity land before editor and chrome polish.
- Non-goal for early milestones: advanced sync/share features before hierarchy, metadata, and interaction rules are stable.

## Milestones

### Milestone 1: Hierarchy and Navigation Parity

- [ ] Match OneNote naming across the app: notebook, section group, section, page, subpage.
- [ ] Create, rename, delete, and reorder notebooks.
- [ ] Create, rename, delete, reorder, collapse, and expand section groups.
- [ ] Create, rename, delete, reorder, and move sections between section groups.
- [ ] Create, rename, delete, reorder, promote, demote, collapse, and expand pages and subpages.
- [ ] Match section-group navigation behavior instead of treating groups as static containers.
- [ ] Match drag/drop rules and feedback for sections and pages.
- [~] Persist notebook hierarchy locally with stable IDs and selection recovery.

Current pass:

- [x] Added a repo spec document to keep the build order explicit.
- [x] Added collapsible section groups.
- [x] Added collapsible subpage stacks in the page pane.
- [x] Added page promote/demote actions.
- [x] Added delete flows and confirmation rules for notebooks, section groups, sections, and pages.
- [x] Added notebook reordering.
- [x] Tightened drag/drop placement rules with before/after indicators and better post-drag selection behavior.
- [ ] Selection behavior parity for collapsed/expanded hierarchy edge cases.
- [ ] Drag/drop rules still need deeper OneNote parity for all edge cases.

### Milestone 2: Visual Parity

- [~] Flatten the chrome and make ribbon density closer to Office.
- [~] Reduce radius, shadows, and decorative gradients where OneNote is flatter.
- [~] Match spacing, borders, inactive states, and typography more tightly.
- [~] Match section tab, page list, and navigation pane emphasis rules.

Current pass:

- [x] Flattened the shell, ribbon, and pane surfaces.
- [x] Reduced radii, shadows, and oversized gradients.
- [x] Made the navigation panes feel more list-like and less card-like.
- [x] Tightened tab treatment and active/inactive selection emphasis.
- [ ] Continue tightening typography, spacing, and exact OneNote control proportions.

### Milestone 3: Editor Parity

- [x] Rich text formatting commands.
- [x] Indentation and list behavior.
- [x] Checkboxes and to-do style interactions.
- [x] Tables.
- [x] Paste behavior from web and Word-like content.
- [x] Images, files, and printout insertion.
- [x] Internal links.
- [x] Better selection and caret handling.
- [x] Template support.

Current pass:

- [x] Wired rich text commands into the Home ribbon for headings, emphasis, color, highlight, links, lists, and paragraph tools.
- [x] Added checklist insertion plus better checklist editing behavior for click, Enter, and empty-item Backspace flows.
- [x] Added table insertion support for everyday note-taking layouts.
- [x] Added sanitized paste handling for HTML, plain text, links, and pasted images.
- [x] Added ribbon-triggered clipboard paste using the async Clipboard API when available.
- [x] Added image insertion, file attachments, and PDF printout cards.
- [x] Added internal page links that navigate back into the notebook hierarchy.
- [x] Preserved editor selection more reliably across ribbon commands and insert flows.
- [x] Added reusable page templates for common note formats.
- [ ] Exact OneNote paste fidelity and full table editing depth remain future polish, but the milestone scope for daily editing is now covered.

### Milestone 4: Data Model and Desktop Features

- [~] Stable notebook storage model with metadata and persisted UI preferences.
- [ ] Attachment and asset storage model.
- [~] Page timestamps, history/version model, and dirty-state handling.
- [~] Global search, recent pages, and sorting modes.
- [~] Keyboard shortcuts, undo/redo, context menus, and autosave indicators.
- [~] App startup persistence and window title/status updates.

Current pass:

- [x] Persisted app-level metadata for recent pages, page sort mode, and search scope.
- [x] Added search scopes for current section, current notebook, and all notebooks.
- [x] Added recent-page tracking and a recent pages panel.
- [x] Added page sorting modes for manual, updated, created, and title order.
- [x] Added clearer dirty/save state feedback plus before-unload protection for unsaved changes.
- [x] Added desktop-style shortcuts for save, search, new page, new notebook, and page promote/demote.
- [x] Updated the window title and status strip to reflect live selection and scope.
- [x] Added hierarchy context menus for notebooks, section groups, sections, and pages.
- [x] Added a central asset registry so inserted images, files, and printouts are stored outside page HTML.
- [x] Added manual page version snapshots with restore support from the History tab.
- [ ] Full multi-step undo/redo parity and richer version browsing still remain future Milestone 4 work.

### Milestone 5: Advanced Features

- [~] Tags and tag panel.
- [~] Outlook-task equivalent workflow.
- [~] Ink/draw layer.
- [~] Audio recording and embedded playback.
- [~] Password-protected sections.
- [ ] Sync/share, only if the product actually requires it.

Current pass:

- [x] Added page tags with visible tag chips in the note header.
- [x] Added a basic follow-up task model with open/done state and optional due date.
- [x] Added a persisted ink board for page-level sketches from the Draw tab.
- [x] Added microphone recording that inserts playable audio notes into the page.
- [x] Added password protection, unlock, relock, and remove-protection flows for sections.
- [ ] Full OneNote-grade freeform ink over the document surface, richer task aggregation, and sync/share are still future Milestone 5 work.

## Core Behavior Checklist

### Structure

- [ ] Notebook CRUD
- [ ] Section group CRUD
- [ ] Section CRUD
- [ ] Page CRUD
- [ ] Subpage CRUD
- [ ] Reorder all hierarchy levels
- [ ] Collapse/expand section groups
- [ ] Collapse/expand subpages

### Editing

- [x] Headings
- [x] Bold, italic, underline
- [x] Bullets and numbering
- [x] Indentation
- [x] Checkboxes
- [x] Tables
- [x] Links
- [x] Highlight
- [x] Images
- [x] Attachments
- [x] Paste handling
- [x] Templates

### Search and Organization

- [x] Search current section
- [x] Search current notebook
- [x] Search all notebooks
- [ ] Search by title, content, and tag
- [x] Recent pages
- [x] Sorting options

### Desktop Behaviors

- [x] Autosave
- [x] Dirty state
- [x] Keyboard shortcuts
- [ ] Undo/redo
- [x] Context menus
- [ ] Drag/drop feedback
- [x] Window title updates

## Implementation Order

1. Finish hierarchy behavior and navigation parity.
2. Tighten visual parity for ribbon, panes, and selection states.
3. Improve the editor for daily note-taking parity.
4. Add attachments, tags, search depth, and shortcuts.
5. Add advanced features only after the earlier layers are stable.
