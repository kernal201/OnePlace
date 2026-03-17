import { escapeHtml } from '../../app/appModel'

export const buildCopilotResponse = (prompt: string, noteText: string) => {
  const sentences = noteText
    .split(/(?<=[.!?])\s+/)
    .map((sentence) => sentence.trim())
    .filter(Boolean)
  const bullets = noteText
    .split(/\n|(?<=\.)\s+/)
    .map((item) => item.replace(/^[-*\u2022]\s*/, '').trim())
    .filter((item) => item.length > 2)
    .slice(0, 6)
  const baseSummary = sentences.slice(0, 3)
  const normalizedPrompt = prompt.toLocaleLowerCase()

  if (normalizedPrompt.includes('organize')) {
    return `
<section class="template-block">
  <h3>Organized Notes</h3>
  <p><strong>Overview:</strong> ${escapeHtml(baseSummary[0] ?? 'Capture the main objective for this page.')}</p>
  <h4>Key Points</h4>
  <ul>${(bullets.length > 0 ? bullets : ['Summarize the important decisions.', 'List the outstanding questions.', 'Capture follow-up work.'])
    .slice(0, 4)
    .map((item) => `<li>${escapeHtml(item)}</li>`)
    .join('')}</ul>
  <h4>Next Steps</h4>
  <ul class="checklist">
    ${(bullets.slice(0, 2).length > 0 ? bullets.slice(0, 2) : ['Review the note and assign owners.', 'Confirm deadlines and dependencies.'])
      .map((item) => `<li><label><input type="checkbox" /> ${escapeHtml(item)}</label></li>`)
      .join('')}
  </ul>
</section>`.trim()
  }

  if (normalizedPrompt.includes('bulleted list')) {
    return `
<section class="template-block">
  <h3>Paragraph Rewrite</h3>
  ${(bullets.length > 0 ? bullets : ['This page would benefit from a more narrative summary.', 'Turn the main ideas into complete thoughts for easier sharing.'])
    .slice(0, 4)
    .map((item) => `<p>${escapeHtml(item.charAt(0).toUpperCase() + item.slice(1))}.</p>`)
    .join('')}
</section>`.trim()
  }

  if (normalizedPrompt.includes('offsite')) {
    return `
<section class="template-block">
  <h3>Team Offsite Plan</h3>
  <p><strong>Goal:</strong> Align the team, make decisions, and leave with owned follow-ups.</p>
  <h4>Day Structure</h4>
  <ul>
    <li>Arrival dinner and context setting.</li>
    <li>Morning strategy workshop and product review.</li>
    <li>Afternoon planning, risk review, and team-building block.</li>
  </ul>
  <h4>Preparation</h4>
  <ul class="checklist">
    <li><label><input type="checkbox" /> Confirm attendees and travel windows.</label></li>
    <li><label><input type="checkbox" /> Draft agenda owners and discussion goals.</label></li>
    <li><label><input type="checkbox" /> Capture decisions and actions in this page.</label></li>
  </ul>
</section>`.trim()
  }

  if (normalizedPrompt.includes('productivity') || normalizedPrompt.includes('manage my time')) {
    return `
<section class="template-block">
  <h3>Productivity Ideas</h3>
  <ul>
    <li>Group similar work into focused blocks and protect them on the calendar.</li>
    <li>Turn open questions into explicit next actions with owners and due dates.</li>
    <li>Use tags to separate urgent work from important but non-urgent work.</li>
    <li>End meetings with a short written summary and assigned follow-ups.</li>
  </ul>
</section>`.trim()
  }

  return `
<section class="template-block">
  <h3>Copilot Draft</h3>
  <p><strong>Prompt:</strong> ${escapeHtml(prompt)}</p>
  <p>${escapeHtml(baseSummary[0] ?? 'Start by summarizing the current page in one sentence.')}</p>
  <ul>${(bullets.length > 0 ? bullets : ['Capture the main point.', 'Call out the risks.', 'Write the next action.'])
    .slice(0, 3)
    .map((item) => `<li>${escapeHtml(item)}</li>`)
    .join('')}</ul>
</section>`.trim()
}
