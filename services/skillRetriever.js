const path = require('path');
const fs = require('fs');

const SKILLS_PATH = path.join(__dirname, '../knowledge/skills.json');
const MAX_SKILLS = 2;

let skillsCache = null;

function loadSkills() {
  if (skillsCache) return skillsCache;
  try {
    const raw = fs.readFileSync(SKILLS_PATH, 'utf-8');
    skillsCache = JSON.parse(raw);
    return skillsCache;
  } catch (err) {
    console.warn('Could not load skills.json:', err.message);
    skillsCache = [];
    return [];
  }
}

/**
 * Find skills relevant to the user message via keyword matching.
 * Returns up to MAX_SKILLS skills, sorted by number of tag hits (most relevant first).
 */
function findRelevantSkills(userMessage) {
  if (!userMessage || typeof userMessage !== 'string') return [];
  const skills = loadSkills();
  if (!Array.isArray(skills) || skills.length === 0) return [];

  const msg = userMessage.toLowerCase();
  const scored = skills
    .map((skill) => {
      const tags = skill.tags || [];
      const hits = tags.filter((tag) => msg.includes(String(tag).toLowerCase()));
      return { ...skill, hitCount: hits.length };
    })
    .filter((s) => s.hitCount > 0)
    .sort((a, b) => b.hitCount - a.hitCount)
    .slice(0, MAX_SKILLS);

  return scored;
}

/**
 * Format matched skills into a string block for injection into the system prompt.
 */
function formatSkillContext(skills) {
  if (!skills || skills.length === 0) return '';
  const blocks = skills.map(
    (s) => `--- [${s.name}] ---\n${s.content || ''}`
  );
  return `\n\nRelevant Excel skill knowledge (use this to guide your response):\n${blocks.join('\n\n')}`;
}

module.exports = { findRelevantSkills, formatSkillContext };
