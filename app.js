/* Service Training Roadmap Tool: page code. Data comes from dataset.xlsx. */
(async function () {
  'use strict';
  /* ======================================================================
     DATA SOURCE
     Everything this page shows comes from dataset.xlsx, which is published
     alongside this page. To update the site, replace dataset.xlsx and
     republish it. Nothing below needs editing for a normal data update.

     Sheets read (matched by name, not position; column order does not matter):
       Programme Index            one row per programme; "Level per Competency" gives the exact
                                  level for each competency (for example "Name: 3; Name: 2")
       Guide - Steps              the "How to use" steps
       Guide - Service Functions  service functions, descriptions, and areas of work
       Guide - Competency Map     which competencies are critical or relevant for each function
       Guide - Proficiency Levels level names, designations, work areas, and role band presets
       Guide - Case Studies       the case studies
     Older workbooks with "Master List" and "Mapping Rules" sheets still work.

     WORDING below only adds friendly descriptions and colours to names that
     already exist in the Excel file. A new name in the Excel file still shows
     up without any change here, with a default colour and description.
  ====================================================================== */
  const DATA_FILE = 'dataset.xlsx';

  const WORDING = {
    roles: {
      'Executive and Senior Executive': 'Typically closest to the public, handling service interactions and cases directly.',
      'Assistant Manager and Manager': 'Typically leads a team or a service area, and handles more complex cases.',
      'Senior Manager and Assistant Director': 'Typically manages a service function or unit, and shapes local practice.',
      'Senior Assistant Director and Deputy Director': 'Typically sets direction for a service area across teams or a department.',
      'Director and Senior Director': 'Typically sets service strategy and standards across a department or agency.',
    },
    levels: {
      1: 'New to this competency, or building basic working knowledge.',
      2: 'Applies this competency independently in routine situations.',
      3: 'Applies this competency confidently across a range of situations.',
      4: 'Applies this competency in complex situations, and guides others.',
      5: 'Shapes policy or practice for this competency across the organisation.',
    },
    competencies: {
      'Customer Service and Case Management': { colour: '#4878D0', desc: 'Handling individual service interactions and cases with the public well, including difficult ones.' },
      'Service Operations Planning and Management': { colour: '#EE854A', desc: 'Running day-to-day service delivery, including workload, resourcing and standard routines.' },
      'Service Quality and Standards': { colour: '#6ACC64', desc: 'Setting and maintaining the standards that service delivery is measured against.' },
      'Service Strategy and Policy Planning': { colour: '#D65F5F', desc: 'Shaping the policies and strategic direction behind how a service is delivered.' },
      'Service Data Management and Analytics': { colour: '#956CB4', desc: 'Using service data and analytics to understand performance and inform decisions.' },
      'Customer Experience and Service Design': { colour: '#DC7EC0', desc: 'Designing or redesigning service experiences around what users actually need.' },
      'Digital Design and Management': { colour: '#8C613C', desc: 'Designing and managing digital service channels and tools.' },
      'Service Tech Innovation and Application': { colour: '#D5BB67', desc: 'Applying new technology, including AI, to improve how a service is delivered.' },
      'Service Workforce Capability Development': { colour: '#82C6E2', desc: 'Building the capability of the people who deliver the service.' },
    },
    values: { 'With Heart': '❤️', 'Know You as One': '🧠', 'Serve You as One': '✨' },
    skills: {
      'Written Communication': { desc: 'For officers who write to the public regularly and want to write with more clarity and empathy.' },
      'Verbal and In-Person Communication': { desc: 'For officers whose work is mostly spoken, face to face, or over the phone, rather than written.' },
      'Emotional Intelligence and Resilience': { desc: 'For officers who deal with emotionally difficult interactions regularly and want to manage that better, for themselves and for the public they serve.' },
      'Data and Analytics': { name: 'Data and Analytics Skills', desc: 'For officers who want to use customer and service data to make better decisions, not just collect it.' },
      'Digital and AI': { name: 'Digital and AI Skills', desc: 'For officers who want practical, hands-on experience with GenAI tools and digital service delivery.' },
      'Handling Difficult and Complex Situations': { desc: 'For officers who regularly deal with upset, complex, or cross-agency cases and want a bigger toolkit for them.' },
      'Customer Experience and Service Design Mindset': { desc: 'For officers who want to think beyond their own interactions to the whole customer journey and service culture around it.' },
    },
    // Add a LEARN playlist link against a skill name here once one exists.
    playlists: {},
  };
  const PALETTE = ['#4878D0', '#EE854A', '#6ACC64', '#D65F5F', '#956CB4', '#DC7EC0', '#8C613C', '#D5BB67', '#82C6E2', '#797979', '#2A9D8F', '#B5838D'];
  const LEVEL_COLOURS = ['#F2B33D', '#F08A2E', '#8DB940', '#2F9BD6', '#D6366A'];

  // Built-in copy for the guide, used only if the guide sheets are missing.
  const GUIDE_DEFAULTS = {
    steps: [
      { n: '1', text: 'Look at your areas of work and their corresponding Service Functions.' },
      { n: '2', text: 'Consider the functional competencies that you would need to have under your Service Function(s) to perform well in your work.' },
      { n: '3', text: "Based on your designation and indicative work areas, refer to the excel titled 'Service Delivery Competencies and Reference Profiles'." },
      { n: '4', text: 'Look at the knowledge and abilities under your respective functional competencies and proficiency levels in the excel.' },
      { n: '5', text: 'Return to the training roadmap and look at the programmes recommended for the selected proficiency level.' },
      { n: 'Tip', text: 'You are welcome to attend a training programme that is pegged at a higher proficiency level if you are looking to upskill.' },
    ],
  };

  /* ---------- small helpers for reading the workbook ---------- */
  function norm(s) { return String(s == null ? '' : s).toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]/g, ''); }
  // Looser key used for matching names to WORDING (ignores a trailing "skills").
  function nameKey(s) { return norm(String(s || '').replace(/\bskills?\b/gi, '')); }
  function clean(v) { return v == null ? '' : String(v).replace(/\r/g, '').trim(); }
  function findWording(map, name) {
    const k = nameKey(name);
    for (const key in map) { if (nameKey(key) === k) return map[key]; }
    return null;
  }
  function initialsId(name, taken) {
    const words = String(name).split(/[\s\-\/]+/).filter(function (w) { return w && !/^(and|of|the|for|&)$/i.test(w); });
    let base = words.map(function (w) { return w[0]; }).join('').toLowerCase().replace(/[^a-z0-9]/g, '') || 'x';
    let id = base, i = 2;
    while (taken.has(id)) { id = base + i; i++; }
    taken.add(id);
    return id;
  }
  function slug(name) { return norm(name).slice(0, 40) || 'item'; }
  function findSheetName(wb, wanted) {
    const w = norm(wanted);
    return wb.SheetNames.find(function (n) { return norm(n) === w; }) ||
      wb.SheetNames.find(function (n) { return norm(n).indexOf(w) !== -1; }) || null;
  }
  function rawRows(wb, wanted) {
    const name = findSheetName(wb, wanted);
    if (!name) return null;
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false, blankrows: true });
  }
  // Rows as objects keyed by normalised header text.
  function tableRows(wb, wanted) {
    const raw = rawRows(wb, wanted);
    if (!raw || !raw.length) return null;
    const head = raw[0].map(norm);
    return raw.slice(1).filter(function (r) { return r.some(function (c) { return clean(c); }); }).map(function (r) {
      const o = {};
      head.forEach(function (h, i) { if (h && !(h in o)) o[h] = clean(r[i]); });
      return o;
    });
  }
  // Reads the first non-empty value among several possible column names.
  function pick(row, names) {
    for (let i = 0; i < names.length; i++) {
      const k = norm(names[i]);
      if (row[k]) return row[k];
    }
    return '';
  }
  function splitList(v) { return clean(v).split(/\s*[,;\n]\s*/).map(clean).filter(Boolean); }
  function splitSemi(v) { return clean(v).split(/\s*[;\n]\s*/).map(clean).filter(Boolean); }
  function parseLevels(v) {
    const out = [];
    const re = /level\s*(\d)\s*[:\-–]?\s*([A-Za-z ]*)/gi;
    let m;
    while ((m = re.exec(String(v || '')))) out.push({ id: Number(m[1]), name: clean(m[2]) });
    if (!out.length) {
      String(v || '').split(/[,;]/).forEach(function (part) {
        const n = parseInt(part, 10);
        if (n >= 1 && n <= 9) out.push({ id: n, name: '' });
      });
    }
    return out;
  }
  function statusKind(s) {
    const k = norm(s);
    if (!k) return 'active';
    if (k.indexOf('discontinu') !== -1) return 'discontinued';
    if (k.indexOf('inhouse') !== -1) return 'inhouse';
    if (k.indexOf('revamp') !== -1) return 'revamping';
    if (k.indexOf('active') !== -1) return 'active';
    return 'other';
  }
  function parseOutcomes(text) {
    const t = clean(text);
    if (!t) return [];
    const lines = t.split(/\n+/).map(function (l) { return l.replace(/^\s*(\d+\s*[.)]|[-•*])\s*/, '').trim(); }).filter(Boolean);
    return lines.map(function (l) { return /[.!?]$/.test(l) ? l : l + '.'; });
  }

  /* ---------- turn a workbook into the catalogue the page uses ---------- */
  function buildCatalogue(wb) {
    const warnings = [];
    const index = tableRows(wb, 'Programme Index') || [];
    const master = tableRows(wb, 'Master List') || [];
    if (!index.length) warnings.push('No "Programme Index" sheet was found, so no active programmes could be listed.');

    // --- Competency map: "Guide - Competency Map", or the older "Mapping Rules" sheet ---
    const compRules = [];   // { name, critical: [fn], relevant: [fn] }
    const levelRoles = {};  // level id -> role band name
    const mapRows = findSheetName(wb, 'Guide - Competency Map') ? (tableRows(wb, 'Guide - Competency Map') || []) : null;
    (mapRows || []).forEach(function (r) {
      const name = pick(r, ['Functional Competency']);
      if (!name) return;
      compRules.push({
        name: name,
        critical: splitSemi(pick(r, ['Critical for (Service Functions, separate with ;)', 'Critical for', 'Critical Competency', 'Critical'])),
        relevant: splitSemi(pick(r, ['Relevant for (Service Functions, separate with ;)', 'Relevant for', 'Relevant Competency', 'Relevant'])),
      });
    });
    const rulesRaw = mapRows ? [] : (rawRows(wb, 'Mapping Rules') || []);
    let mode = null;
    rulesRaw.forEach(function (r) {
      const a = clean(r[0]), b = clean(r[1]), c = clean(r[2]);
      const ka = norm(a);
      if (ka === 'functionalcompetency' || ka === 'functionalcompetencies') { mode = 'comp'; return; }
      if (ka === 'competencylevels' || ka === 'competencylevel' || ka === 'proficiencylevels') { mode = 'level'; return; }
      if (!a) return;
      if (mode === 'comp') compRules.push({ name: a, critical: splitSemi(b), relevant: splitSemi(c) });
      else if (mode === 'level') { const lv = parseLevels(a)[0]; if (lv) levelRoles[lv.id] = b; }
    });

    // --- Guide sheets ---
    const guideStepsRows = tableRows(wb, 'Guide - Steps');
    const guideFnRows = tableRows(wb, 'Guide - Service Functions');
    const guideLevelRows = tableRows(wb, 'Guide - Proficiency Levels');

    // --- Proficiency levels ---
    const levelMap = {};
    function ensureLevel(id, name) {
      if (!levelMap[id]) levelMap[id] = { id: id, name: name || ('Level ' + id), desc: WORDING.levels[id] || '', colour: LEVEL_COLOURS[(id - 1) % LEVEL_COLOURS.length], designation: '', workAreas: '' };
      else if (name && /^Level \d$/.test(levelMap[id].name)) levelMap[id].name = name;
      return levelMap[id];
    }
    (guideLevelRows || []).forEach(function (r) {
      const id = parseInt(pick(r, ['Level']), 10);
      if (!id) return;
      const l = ensureLevel(id, pick(r, ['Name']));
      l.designation = pick(r, ['Typical designation', 'Designation']);
      l.workAreas = pick(r, ['Typical work areas', 'Work areas']);
      l.bandName = pick(r, ['Band preset name', 'Band name']);
      l.bandDesc = pick(r, ['Band preset description', 'Band description']);
      const col = pick(r, ['Colour (hex)', 'Colour', 'Color']);
      if (/^#?[0-9a-f]{6}$/i.test(col)) l.colour = col.charAt(0) === '#' ? col : '#' + col;
    });
    Object.keys(levelRoles).forEach(function (id) { const l = ensureLevel(Number(id)); if (!l.designation) l.designation = levelRoles[id]; });
    index.concat(master).forEach(function (r) {
      parseLevels(pick(r, ['Competency Levels Mapped', 'Competency Level', 'Competency Levels'])).forEach(function (lv) { ensureLevel(lv.id, lv.name); });
    });
    const proficiencyLevels = Object.keys(levelMap).map(Number).sort(function (a, b) { return a - b; }).map(function (id) { return levelMap[id]; });

    // --- Role bands: ordered by level, then any extra names in the data ---
    const roleBands = [];
    const roleByName = {};
    function ensureRole(name) {
      const k = norm(name);
      if (!k) return null;
      if (!roleByName[k]) {
        const r = { id: slug(name), name: name, desc: findWording(WORDING.roles, name) || '' };
        roleByName[k] = r; roleBands.push(r);
      }
      return roleByName[k];
    }
    proficiencyLevels.forEach(function (l) { if (l.designation) ensureRole(l.designation); });
    index.forEach(function (r) { splitList(pick(r, ['Recommended For', 'Recommendation'])).forEach(ensureRole); });

    // --- Functional competencies ---
    const takenIds = new Set();
    const competencies = [];
    const compByName = {};
    function ensureComp(name) {
      const k = nameKey(name);
      if (!k) return null;
      if (!compByName[k]) {
        const w = findWording(WORDING.competencies, name) || {};
        const c = { id: initialsId(name, takenIds), name: name, colour: w.colour || PALETTE[competencies.length % PALETTE.length], desc: w.desc || '' };
        compByName[k] = c; competencies.push(c);
      }
      return compByName[k];
    }
    compRules.forEach(function (r) { ensureComp(r.name); });
    index.concat(master).forEach(function (r) { splitList(pick(r, ['Functional Competencies Mapped', 'Functional Competency'])).forEach(ensureComp); });

    // --- Service values ---
    const serviceValues = [];
    const valueByName = {};
    index.forEach(function (r) {
      splitList(pick(r, ['Service Values Mapped', 'Service Value'])).forEach(function (name) {
        const k = norm(name);
        if (!valueByName[k]) {
          const v = { id: slug(name), name: name, symbol: findWording(WORDING.values, name) || '◆' };
          valueByName[k] = v; serviceValues.push(v);
        }
      });
    });

    // --- Skill areas: friendly order first, then anything new in the data ---
    const skillAreas = [];
    const skillByKey = {};
    function ensureSkill(name) {
      const k = nameKey(name);
      if (!k) return null;
      if (!skillByKey[k]) {
        const w = findWording(WORDING.skills, name) || {};
        const s = { id: slug(name), name: w.name || name, desc: w.desc || ('Programmes that build ' + name.toLowerCase() + '.'), excelName: name, order: 999 };
        skillByKey[k] = s; skillAreas.push(s);
      }
      return skillByKey[k];
    }

    // --- Programmes ---
    const programmes = [];
    const comingSoonProgrammes = [];
    const discontinuedMap = new Map();
    const inHouseMap = new Map();
    const seenIds = new Set();
    index.forEach(function (r) {
      const code = pick(r, ['Code', 'Programme Code']);
      const title = pick(r, ['Programme', 'Programme Title', 'Title']);
      if (!title) return;
      const kind = statusKind(pick(r, ['Status']));
      const levels = parseLevels(pick(r, ['Competency Levels Mapped', 'Competency Level'])).map(function (l) { return l.id; });
      const modality = pick(r, ['Modality']) || 'Not stated';
      const duration = pick(r, ['Duration']) || 'TBC';
      if (kind === 'discontinued' || kind === 'inhouse') {
        const target = kind === 'discontinued' ? discontinuedMap : inHouseMap;
        const key = code + '|' + title;
        if (!target.has(key)) target.set(key, { code: code || 'TBC', title: title, modality: modality, duration: duration, levels: [], competencyIds: [] });
        return;
      }
      if (!levels.length || kind === 'other') {
        comingSoonProgrammes.push({ code: code || 'TBC', title: title, modality: modality });
        return;
      }
      let id = code || slug(title);
      while (seenIds.has(id)) id = id + '-2';
      seenIds.add(id);
      const listing = pick(r, ['Listings', 'Listing', 'URL', 'Link']);
      const skillNames = splitList(pick(r, ['Skills', 'Skill', 'Skill Areas']));
      programmes.push({
        id: id,
        title: title,
        modality: modality,
        duration: duration,
        status: kind === 'revamping' ? 'Revamping' : 'Active',
        levels: Array.from(new Set(levels)).sort(function (a, b) { return a - b; }),
        valueIds: splitList(pick(r, ['Service Values Mapped', 'Service Value'])).map(function (n) { return valueByName[norm(n)].id; }),
        competencyIds: splitList(pick(r, ['Functional Competencies Mapped', 'Functional Competency'])).map(function (n) { return ensureComp(n).id; }),
        roleIds: splitList(pick(r, ['Recommended For', 'Recommendation'])).map(function (n) { return ensureRole(n).id; }),
        criticalRaw: pick(r, ['Critical Competencies Mapped', 'Critical Competency']),
        relevantRaw: pick(r, ['Relevant Competencies Mapped', 'Relevant Competency']),
        remarks: pick(r, ['Remarks']),
        url: /^https?:\/\//i.test(listing) ? listing : '',
        skillIds: skillNames.map(function (n) { return ensureSkill(n).id; }),
        overview: pick(r, ['Description', 'Overview']),
        learningOutcomes: parseOutcomes(pick(r, ['Learning Outcomes'])),
        outline: pick(r, ['Outline']),
        _levelPerComp: pick(r, ['Level per Competency', 'Levels per Competency']),
      });
    });

    master.forEach(function (r) {
      const kind = statusKind(pick(r, ['Status']));
      if (kind !== 'discontinued' && kind !== 'inhouse') return;
      const code = pick(r, ['Code']) || 'TBC';
      const title = pick(r, ['Programme', 'Title']);
      if (!title) return;
      const target = kind === 'discontinued' ? discontinuedMap : inHouseMap;
      const key = code + '|' + title;
      if (!target.has(key)) target.set(key, { code: code, title: title, modality: pick(r, ['Modality']), duration: pick(r, ['Duration']) || 'TBC', levels: [], competencyIds: [] });
      const item = target.get(key);
      parseLevels(pick(r, ['Competency Level', 'Competency Levels Mapped'])).forEach(function (l) { if (item.levels.indexOf(l.id) === -1) item.levels.push(l.id); });
      splitList(pick(r, ['Functional Competency', 'Functional Competencies Mapped'])).forEach(function (n) { const c = ensureComp(n); if (item.competencyIds.indexOf(c.id) === -1) item.competencyIds.push(c.id); });
    });
    // Which level each programme is pegged at for each competency. The Master
    // List has one row per programme and competency, so this is more exact
    // than crossing the merged lists in the Programme Index.
    const compLevelsByCode = {};
    master.forEach(function (r) {
      const kind = statusKind(pick(r, ['Status']));
      if (kind !== 'active' && kind !== 'revamping') return;
      const code = pick(r, ['Code']);
      if (!code) return;
      const lv = parseLevels(pick(r, ['Competency Level', 'Competency Levels Mapped']));
      splitList(pick(r, ['Functional Competency', 'Functional Competencies Mapped'])).forEach(function (n) {
        const c = ensureComp(n);
        compLevelsByCode[code] = compLevelsByCode[code] || {};
        const set = compLevelsByCode[code][c.id] = compLevelsByCode[code][c.id] || [];
        lv.forEach(function (l) { if (set.indexOf(l.id) === -1) set.push(l.id); });
      });
    });
    programmes.forEach(function (p) {
      const fromMaster = compLevelsByCode[p.id];
      const fromIndex = {};
      // "Level per Competency" in the Programme Index, for example "Name: 3; Name: 2, 3"
      splitSemi(p._levelPerComp || '').forEach(function (part) {
        const m = part.match(/^(.*?):\s*([\d,\s]+)$/);
        if (!m) return;
        const c = compByName[nameKey(m[1])];
        if (!c) { warnings.push(p.title + ': competency "' + clean(m[1]) + '" in Level per Competency was not recognised.'); return; }
        fromIndex[c.id] = m[2].split(',').map(function (x) { return parseInt(x, 10); }).filter(function (x) { return x >= 1; });
      });
      p.compLevels = {};
      p.competencyIds.forEach(function (cid) {
        const exact = (fromIndex[cid] && fromIndex[cid].length) ? fromIndex[cid] : (fromMaster && fromMaster[cid] && fromMaster[cid].length ? fromMaster[cid] : null);
        p.compLevels[cid] = exact ? exact.slice().sort() : p.levels.slice();
      });
      delete p._levelPerComp;
    });

    const discontinuedProgrammes = Array.from(discontinuedMap.values()).map(function (p) { p.levels.sort(); return p; });
    const inHouseProgrammes = Array.from(inHouseMap.values());

    // Order skill areas: the friendly list first, in its own order.
    const skillOrder = Object.keys(WORDING.skills).map(nameKey);
    skillAreas.forEach(function (s) { const i = skillOrder.indexOf(nameKey(s.excelName)); s.order = i === -1 ? 999 : i; });
    skillAreas.sort(function (a, b) { return a.order - b.order; });

    function joinNice(list) { return list.length > 1 ? list.slice(0, -1).join(', ') + ' and ' + list[list.length - 1] : (list[0] || ''); }
    // Preset pathways by role band: every programme recommended for that band.
    const bandPresets = roleBands.map(function (rb) {
      const lvl = proficiencyLevels.find(function (l) { return norm(l.designation) === norm(rb.name); });
      const members = programmes.filter(function (p) { return p.roleIds.indexOf(rb.id) !== -1; });
      if (!members.length) return null;
      const mods = Array.from(new Set(members.map(function (p) { return p.modality; }))).sort();
      return {
        id: 'band-' + rb.id,
        category: 'band',
        role: rb.id,
        level: lvl ? lvl.id : null,
        title: (lvl && lvl.bandName) || rb.name,
        subtitle: rb.name,
        desc: (lvl && lvl.bandDesc) || rb.desc || ('Programmes recommended for ' + rb.name + '.'),
        skillFocus: lvl ? 'Proficiency Level ' + lvl.id + ' (' + lvl.name + ')' : rb.name,
        modality: joinNice(mods),
        playlistUrl: findWording(WORDING.playlists, (lvl && lvl.bandName) || rb.name) || '',
      };
    }).filter(Boolean);

    // Preset pathways: one per skill area that has at least one programme.
    const skillPresets = skillAreas.map(function (s) {
      const members = programmes.filter(function (p) { return p.skillIds.indexOf(s.id) !== -1; });
      if (!members.length) return null;
      const mods = Array.from(new Set(members.map(function (p) { return p.modality; }))).sort();
      return {
        id: 'skill-' + s.id,
        category: 'skill',
        skillAreaId: s.id,
        title: s.name,
        desc: s.desc,
        skillFocus: s.name,
        modality: joinNice(mods),
        playlistUrl: findWording(WORDING.playlists, s.excelName) || '',
      };
    }).filter(Boolean);
    const presetPathways = bandPresets.concat(skillPresets);

    // --- Service functions for the guide ---
    const serviceFunctions = [];
    const fnByKey = {};
    function ensureFn(name) {
      const k = norm(name);
      if (!k) return null;
      if (!fnByKey[k]) { const f = { id: 'fn-' + slug(name), name: name, desc: '', areas: [] }; fnByKey[k] = f; serviceFunctions.push(f); }
      return fnByKey[k];
    }
    (guideFnRows || []).forEach(function (r) {
      const f = ensureFn(pick(r, ['Service Function']));
      if (!f) return;
      f.desc = pick(r, ['Description']);
      f.areas = clean(pick(r, ['Areas of Work (one per line)', 'Areas of Work'])).split(/\n+/).map(function (a) { return a.replace(/^\s*(\d+\s*[.)]|[-•*])\s*/, '').trim(); }).filter(Boolean);
    });
    compRules.forEach(function (r) { r.critical.concat(r.relevant).forEach(ensureFn); });
    if (!guideFnRows) warnings.push('No "Guide - Service Functions" sheet was found, so service function descriptions are not shown.');

    const competencyRules = compRules.map(function (r) {
      return {
        compId: ensureComp(r.name).id,
        critical: r.critical.map(function (n) { return ensureFn(n).id; }),
        relevant: r.relevant.map(function (n) { return ensureFn(n).id; }),
      };
    });

    const guideSteps = guideStepsRows && guideStepsRows.length
      ? guideStepsRows.map(function (r) { return { n: pick(r, ['Step']), text: pick(r, ['Instruction', 'Text']) }; }).filter(function (s) { return s.text; })
      : GUIDE_DEFAULTS.steps;

    const caseRows = tableRows(wb, 'Guide - Case Studies') || [];
    const caseStudies = caseRows.map(function (r, i) {
      const needs = [];
      splitSemi(pick(r, ['Competencies and required levels (Competency: level; ...)', 'Competencies and required levels', 'Competencies'])).forEach(function (part) {
        const m = part.match(/^(.*?)[:\-\u2013]\s*(?:level\s*)?(\d)\s*$/i);
        if (!m) return;
        const comp = compByName[nameKey(m[1])];
        if (!comp) { warnings.push('Case study ' + (i + 1) + ': competency "' + clean(m[1]) + '" was not recognised.'); return; }
        needs.push({ compId: comp.id, level: Number(m[2]) });
      });
      return {
        n: pick(r, ['Case']) || String(i + 1),
        title: pick(r, ['Title']),
        persona: pick(r, ['Persona']),
        designation: pick(r, ['Designation']),
        story: pick(r, ['Storyline', 'Story']),
        functions: splitSemi(pick(r, ['Service Functions (separate with ;)', 'Service Functions'])).map(function (n) { return ensureFn(n); }).filter(Boolean).map(function (f) { return f.id; }),
        needs: needs,
      };
    }).filter(function (c) { return c.title && c.needs.length; });

    return {
      caseStudies: caseStudies,
      roleBands: roleBands, proficiencyLevels: proficiencyLevels, competencies: competencies, serviceValues: serviceValues,
      skillAreas: skillAreas, programmes: programmes, discontinuedProgrammes: discontinuedProgrammes,
      inHouseProgrammes: inHouseProgrammes, comingSoonProgrammes: comingSoonProgrammes, presetPathways: presetPathways,
      serviceFunctions: serviceFunctions, competencyRules: competencyRules, guideSteps: guideSteps, warnings: warnings,
    };
  }

  // Where the data comes from, in order:
  //   1. dataset.xlsx next to this page (normal web hosting: replace the file to update the site)
  //   2. dataset.js, a wrapped copy of the same Excel file (used where .xlsx files cannot be served,
  //      such as the claude.ai artifact, or when the page is opened straight from a folder)
  let dataOrigin = '';
  async function fetchWorkbook() {
    if (/^https?:$/.test(location.protocol)) {
      try {
        const res = await fetch(DATA_FILE, { cache: 'no-store' });
        if (res.ok) {
          const buf = await res.arrayBuffer();
          const b = new Uint8Array(buf.slice(0, 2));
          if (b[0] === 0x50 && b[1] === 0x4B) { // a real .xlsx file starts with "PK"
            dataOrigin = DATA_FILE;
            return XLSX.read(buf, { type: 'array' });
          }
        }
      } catch (e) { /* fall through to dataset.js */ }
    }
    if (window.DATASET_XLSX && window.DATASET_XLSX.base64) {
      dataOrigin = (window.DATASET_XLSX.name || DATA_FILE) + (window.DATASET_XLSX.updated ? ', updated ' + window.DATASET_XLSX.updated : '');
      return XLSX.read(window.DATASET_XLSX.base64, { type: 'base64' });
    }
    if (location.protocol === 'file:') {
      throw new Error('This page was opened straight from a folder, so the browser will not let it read ' + DATA_FILE + '. Please open it through its web address instead.');
    }
    throw new Error('Could not load ' + DATA_FILE + '. Please check that it is in the same folder as this page on the server.');
  }


  /* ======================================================================
     LIVE CATALOGUE — filled from the Excel file by applyCatalogue()
  ====================================================================== */
  let roleBands = [], proficiencyLevels = [], competencies = [], serviceValues = [], skillAreas = [];
  let programmes = [], discontinuedProgrammes = [], inHouseProgrammes = [], comingSoonProgrammes = [], presetPathways = [];
  let serviceFunctions = [], competencyRules = [], guideSteps = [], dataWarnings = [], caseStudies = [];
  let modalityOptions = [];
  let compById = {}, levelById = {}, roleById = {}, valueById = {}, skillAreaById = {}, progById = {}, fnById = {};
  let dataSourceLabel = '';
  const LEARN_HOME_URL = 'https://www.learn.gov.sg/Learning/Home';
  // Shown wherever real course content (overview/outcomes/outline) is
  // missing for a programme, instead of inventing a stand-in summary.
  const NO_DATA_MESSAGE = 'Data not available currently. Please check with the programme owner.';

  function applyCatalogue(cat, label) {
    roleBands = cat.roleBands; proficiencyLevels = cat.proficiencyLevels; competencies = cat.competencies;
    serviceValues = cat.serviceValues; skillAreas = cat.skillAreas; programmes = cat.programmes;
    discontinuedProgrammes = cat.discontinuedProgrammes; inHouseProgrammes = cat.inHouseProgrammes;
    comingSoonProgrammes = cat.comingSoonProgrammes; presetPathways = cat.presetPathways;
    serviceFunctions = cat.serviceFunctions; competencyRules = cat.competencyRules; guideSteps = cat.guideSteps;
    dataWarnings = cat.warnings; caseStudies = cat.caseStudies;
    modalityOptions = Array.from(new Set(programmes.map(function (p) { return p.modality; }))).sort();
    compById = Object.fromEntries(competencies.map(function (c) { return [c.id, c]; }));
    levelById = Object.fromEntries(proficiencyLevels.map(function (l) { return [l.id, l]; }));
    roleById = Object.fromEntries(roleBands.map(function (r) { return [r.id, r]; }));
    valueById = Object.fromEntries(serviceValues.map(function (v) { return [v.id, v]; }));
    skillAreaById = Object.fromEntries(skillAreas.map(function (s) { return [s.id, s]; }));
    progById = Object.fromEntries(programmes.map(function (p) { return [p.id, p]; }));
    fnById = Object.fromEntries(serviceFunctions.map(function (f) { return [f.id, f]; }));
    dataSourceLabel = label;
  }

  function levelLabel(id) { return levelById[id] ? levelById[id].name : id; }
  function levelCode(ids) { return ids.map(function (i) { return 'L' + i; }).join('–'); }
  function roleNames(ids) { return ids.map(function (r) { return roleById[r] ? roleById[r].name : r; }).join(', '); }
  function compNames(ids) { return ids.map(function (c) { return compById[c] ? compById[c].name : c; }).join(', '); }

  // A rationale for a skill-derived step, where there's no single role or
  // current level to anchor against (a skill preset spans several role
  // bands and levels by design) — grounded in the programme's own mapped
  // competency and level, plus a concrete learning-outcome hook where one
  // exists, same as buildRationale above.
  function buildSkillNote(programme) {
    let sentence = 'Maps to ' + compNames(programme.competencyIds) + ' at ' + levelCode(programme.levels) + ' level.';
    if (programme.status === 'Revamping') sentence += ' This programme is currently being revamped, so details may change.';
    if (programme.learningOutcomes && programme.learningOutcomes.length) {
      const firstOutcome = programme.learningOutcomes[0].replace(/\.$/, '');
      sentence += ' By the end, you will be able to ' + firstOutcome.charAt(0).toLowerCase() + firstOutcome.slice(1) + '.';
    }
    return sentence;
  }

  // Every active programme tagged with a given skill area, freshly derived
  // from the live catalogue's skillIds field each time — never a stored
  // list, so a catalogue update that changes which programmes carry a
  // skill tag is reflected here automatically.
  function programmesForSkill(skillAreaId) {
    return programmes.filter(function (p) { return p.skillIds && p.skillIds.includes(skillAreaId); });
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // Turns raw multi-paragraph text (blank line = new paragraph, single line
  // break kept as a line break) into safe, readable HTML.
  function renderMultiParagraph(text) {
    if (!text) return '';
    return text.split(/\n\s*\n/)
      .map(function (para) { return para.trim(); })
      .filter(Boolean)
      .map(function (para) { return '<p>' + escapeHtml(para).replace(/\n/g, '<br>') + '</p>'; })
      .join('');
  }

  // Same paragraph-block splitting as above, but a block whose lines start
  // with "-" (the convention used throughout the real outline data) renders
  // as an actual bullet list instead of plain "-text" lines joined by
  // <br> — a block can mix a plain header line followed by bullet lines,
  // which is the common real shape ("CARP Framework\n-Overview\n-Practice").
  function renderOutline(text) {
    if (!text) return '';
    return text.split(/\n\s*\n/)
      .map(function (block) { return block.trim(); })
      .filter(Boolean)
      .map(function (block) {
        const lines = block.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
        const isBullet = function (l) { return /^-\s*/.test(l); };
        if (!lines.some(isBullet)) {
          return '<p>' + escapeHtml(block).replace(/\n/g, '<br>') + '</p>';
        }
        let html = '';
        let currentBullets = [];
        function flushBullets() {
          if (currentBullets.length) {
            html += '<ul class="outline-list">' + currentBullets.map(function (l) {
              return '<li>' + escapeHtml(l.replace(/^-\s*/, '')) + '</li>';
            }).join('') + '</ul>';
            currentBullets = [];
          }
        }
        lines.forEach(function (l) {
          if (isBullet(l)) { currentBullets.push(l); }
          else { flushBullets(); html += '<p>' + escapeHtml(l) + '</p>'; }
        });
        flushBullets();
        return html;
      })
      .join('');
  }

  /* ======================================================================
     STATE
  ====================================================================== */
  const state = {
    theme: 'light',
    view: 'home',
    selectedPreset: null,
    browseFilters: { comp: new Set(), level: new Set(), role: new Set(), modality: new Set(), value: new Set() },
    browseSearch: '',
    browseSort: 'competency',
    matrixFilters: { search: '', comp: '', level: '', modality: '', role: '' },
    pathway: new Map(), // programme id -> rationale string or null
    pathwayName: '',
    pathwayOrg: '',
  };

  /* ======================================================================
     PATHWAY (persistent tray, available from every view)
  ====================================================================== */
  const pathwayCountEl = document.getElementById('pathway-count');
  const pathwayOverlay = document.getElementById('pathway-overlay');
  const pathwayBody = document.getElementById('pathway-body');
  const presetOverlay = document.getElementById('preset-overlay');

  function togglePathway(id, rationale) {
    if (state.pathway.has(id)) state.pathway.delete(id);
    else state.pathway.set(id, rationale || null);
    syncAfterPathwayChange();
  }
  function removeFromPathway(id) { state.pathway.delete(id); syncAfterPathwayChange(); }

  function syncAfterPathwayChange() {
    pathwayCountEl.textContent = String(state.pathway.size);
    if (state.view === 'home') renderBrowseResults();
    if (state.view === 'cases') { const op = document.querySelector('.case-stretch'); const wasOpen = op && op.open; renderCases(); const np = document.querySelector('.case-stretch'); if (np && wasOpen) np.open = true; }
    if (presetOverlay.classList.contains('open')) renderPresetResults();
    if (pathwayOverlay.classList.contains('open')) renderPathwayDrawer();
    if (modalOverlay.classList.contains('open') && currentModalProgId) refreshModalAddButton();
  }

  function addBtnHtml(id, extraClass, rationale) {
    const added = state.pathway.has(id);
    const rat = rationale ? " data-rationale=\"" + rationale.replace(/"/g, '&quot;') + "\"" : '';
    return '<button class="btn-add ' + (extraClass || '') + (added ? ' added' : '') + '" data-toggle-pathway="' + id + '"' + rat + '>' +
      (added ? '\u2713 In your pathway' : '+ Add to pathway') + '</button>';
  }

  document.addEventListener('click', function (e) {
    const btn = e.target.closest('[data-toggle-pathway]');
    if (btn) {
      togglePathway(btn.dataset.togglePathway, btn.dataset.rationale || null);
    }
  });

  function openPathway() {
    presetOverlay.classList.remove('open');
    if (typeof closeModal === 'function') closeModal();
    renderPathwayDrawer();
    pathwayOverlay.classList.add('open');
    const closeBtn = pathwayOverlay.querySelector('[data-close-pathway]');
    if (closeBtn) closeBtn.focus();
  }
  document.getElementById('pathway-open').addEventListener('click', openPathway);
  document.querySelector('[data-close-pathway]').addEventListener('click', function () { pathwayOverlay.classList.remove('open'); });
  pathwayOverlay.addEventListener('click', function (e) { if (e.target === pathwayOverlay) pathwayOverlay.classList.remove('open'); });
  document.querySelector('[data-close-preset]').addEventListener('click', function () { presetOverlay.classList.remove('open'); });
  presetOverlay.addEventListener('click', function (e) { if (e.target === presetOverlay) presetOverlay.classList.remove('open'); });

  const pathwayOpen = new Set(); // programmes expanded in the pathway window
  function renderPathwayDrawer() {
    if (state.pathway.size === 0) {
      pathwayBody.innerHTML = '<p class="empty-state" style="padding:20px 0;">Nothing added yet. Add programmes from the catalogue on the home page, or from a preset pathway.</p>';
      return;
    }
    const rows = Array.from(state.pathway.keys()).filter(function (id) { return progById[id]; }).map(function (id, i) {
      const p = progById[id];
      const open = pathwayOpen.has(id);
      const summary = p.overview ? (p.overview.length > 260 ? p.overview.slice(0, 260).replace(/\s+\S*$/, '') + '\u2026' : p.overview) : NO_DATA_MESSAGE;
      return '<div class="pw-item' + (open ? ' open' : '') + '">' +
        '<div class="pw-head-row">' +
          '<button class="pw-head" data-pw-toggle="' + id + '" aria-expanded="' + open + '">' +
            '<span class="pw-num">' + (i + 1) + '</span>' +
            '<span class="pw-main"><span class="p-title">' + escapeHtml(p.title) + '</span>' +
            '<span class="p-meta">' + levelCode(p.levels) + ' &middot; ' + escapeHtml(p.modality) + ' &middot; ' + escapeHtml(p.duration) + '</span></span>' +
            '<span class="pw-chev" aria-hidden="true">\u25BC</span>' +
          '</button>' +
          '<button class="p-remove" data-remove-pathway="' + id + '" aria-label="Remove ' + escapeHtml(p.title) + '">&times;</button>' +
        '</div>' +
        '<div class="pw-body">' +
          '<p class="pw-summary">' + escapeHtml(summary) + '</p>' +
          '<div class="pw-facts">' +
            '<div><span class="d-label">Competencies</span><span>' + escapeHtml(compNames(p.competencyIds)) + '</span></div>' +
            '<div><span class="d-label">Recommended for</span><span>' + escapeHtml(roleNames(p.roleIds)) + '</span></div>' +
          '</div>' +
          '<div class="btn-row">' +
            (p.url ? '<a class="btn btn-secondary btn-sm" href="' + p.url + '" target="_blank" rel="noopener">Open on LEARN</a>' : '<span class="pill-nomination">' + escapeHtml(p.remarks || 'By nomination only') + '</span>') +
            '<button class="btn btn-secondary btn-sm" data-pw-detail="' + id + '">More details</button>' +
          '</div>' +
        '</div>' +
      '</div>';
    }).join('');
    pathwayBody.innerHTML =
      '<p class="step-sub" style="margin-bottom:14px;">' + state.pathway.size + ' programme' + (state.pathway.size === 1 ? '' : 's') + ' in your pathway.</p>' +
      '<div class="pathway-list">' + rows + '</div>' +
      '<p class="pw-fields-note">Add your details. They are printed on your PDF, and are not saved or sent anywhere.</p>' +
      '<div class="pw-fields">' +
        '<label class="pw-field"><span>Name</span><input type="text" id="pathway-name-input" class="pathway-name-field" autocomplete="name" placeholder="e.g. Tan Mei Ling" value="' + escapeHtml(state.pathwayName).replace(/"/g, '&quot;') + '"></label>' +
        '<label class="pw-field"><span>Organisation</span><input type="text" id="pathway-org-input" class="pathway-name-field" autocomplete="organization" placeholder="e.g. your agency or department" value="' + escapeHtml(state.pathwayOrg).replace(/"/g, '&quot;') + '"></label>' +
      '</div>' +
      '<div class="btn-row"><button class="btn btn-primary" id="pathway-export">Export as PDF</button><button class="btn btn-secondary" id="pathway-clear">Clear all</button></div>';

    pathwayBody.querySelectorAll('[data-pw-toggle]').forEach(function (b) {
      b.addEventListener('click', function () {
        const id = b.dataset.pwToggle;
        if (pathwayOpen.has(id)) pathwayOpen.delete(id); else pathwayOpen.add(id);
        const item = b.closest('.pw-item');
        item.classList.toggle('open', pathwayOpen.has(id));
        b.setAttribute('aria-expanded', String(pathwayOpen.has(id)));
      });
    });
    pathwayBody.querySelectorAll('[data-pw-detail]').forEach(function (b) {
      b.addEventListener('click', function () { pathwayOverlay.classList.remove('open'); openModal(progById[b.dataset.pwDetail]); });
    });
    pathwayBody.querySelectorAll('[data-remove-pathway]').forEach(function (b) {
      b.addEventListener('click', function () { removeFromPathway(b.dataset.removePathway); });
    });
    document.getElementById('pathway-name-input').addEventListener('input', function (e) { state.pathwayName = e.target.value; });
    document.getElementById('pathway-org-input').addEventListener('input', function (e) { state.pathwayOrg = e.target.value; });
    document.getElementById('pathway-export').addEventListener('click', exportPathwayPdf);
    document.getElementById('pathway-clear').addEventListener('click', function () { state.pathway.clear(); syncAfterPathwayChange(); });
  }

  /* ======================================================================
     PDF EXPORT — styled to match the website (navy ink, dark grey bars,
     light grey panels). Built with jsPDF in the browser. In a claude.ai
     artifact the file is saved through the platform's "downloads"
     capability; elsewhere it falls back to a normal browser download.
  ====================================================================== */
  const PDF_COLOURS = {
    ink: [18, 32, 58],        // --ink
    inkSoft: [75, 88, 114],   // --ink-soft
    inkFaint: [124, 135, 154],// --ink-faint
    bar: [58, 58, 58],        // dark grey used for table headers
    panel: [236, 238, 233],   // --surface-alt
    border: [217, 221, 212],  // --border
    link: [46, 92, 170],
  };
  function hexToRgb(hex) {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex || '');
    return m ? [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)] : PDF_COLOURS.inkFaint;
  }
  function shortText(text, max) {
    const t = String(text || '').replace(/\s+/g, ' ').trim();
    if (t.length <= max) return t;
    return t.slice(0, max).replace(/\s+\S*$/, '') + '…';
  }
  // jsPDF's built-in fonts cover Western European characters only.
  function pdfSafe(s) {
    return String(s == null ? '' : s)
      .replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
      .replace(/[^\x09\x0A\x0D\x20-\x7E -ÿ–—•…€]/g, '');
  }

  async function exportPathwayPdf() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      // the main copy did not load: try the backup CDN once
      try { await loadFirst(JSPDF_SOURCES, jspdfReady, 'the PDF tool (jsPDF)'); } catch (e) { /* handled below */ }
    }
    if (!window.jspdf || !window.jspdf.jsPDF) {
      showToast('Could not load the PDF export tool. Please check your connection and try again.');
      return;
    }
    const ids = Array.from(state.pathway.keys()).filter(function (id) { return progById[id]; });
    if (!ids.length) { showToast('Add at least one programme to your pathway first.'); return; }

    const doc = new window.jspdf.jsPDF({ unit: 'pt', format: 'a4' });
    doc.setLineHeightFactor(1.3); // keeps drawn line spacing in step with the height sums below
    const W = doc.internal.pageSize.getWidth();
    const H = doc.internal.pageSize.getHeight();
    const M = 44;                 // page margin
    const CW = W - M * 2;         // content width
    const FOOT = 46;              // space kept for the footer
    let y = 0;
    const now = new Date();
    const dateText = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

    function setText(rgb, size, bold) {
      doc.setTextColor(rgb[0], rgb[1], rgb[2]);
      doc.setFontSize(size);
      doc.setFont('helvetica', bold ? 'bold' : 'normal');
    }
    function lines(text, width, size, bold) {
      doc.setFontSize(size); doc.setFont('helvetica', bold ? 'bold' : 'normal');
      return doc.splitTextToSize(pdfSafe(text), width);
    }
    function newPage() {
      doc.addPage();
      // slim header bar on follow-on pages
      doc.setFillColor.apply(doc, PDF_COLOURS.bar);
      doc.rect(0, 0, W, 30, 'F');
      setText([255, 255, 255], 9, true);
      doc.text('Service Training Roadmap Tool', M, 19);
      setText([210, 210, 210], 9, false);
      doc.text('My training pathway', W - M, 19, { align: 'right' });
      y = 54;
    }
    function ensure(h) { if (y + h > H - FOOT) newPage(); }

    // ---- Header band ----
    doc.setFillColor.apply(doc, PDF_COLOURS.bar);
    doc.rect(0, 0, W, 92, 'F');
    // brand mark: two overlapping squares, as on the site
    doc.setFillColor(255, 255, 255); doc.rect(M, 26, 14, 14, 'F');
    doc.setDrawColor(255, 255, 255); doc.setLineWidth(1.5); doc.rect(M + 9, 35, 14, 14, 'S');
    setText([200, 200, 200], 9.5, false);
    doc.text('Service Training Roadmap Tool', M + 34, 36);
    setText([255, 255, 255], 22, true);
    doc.text('My training pathway', M + 34, 62);
    y = 118;

    // ---- Details ----
    const name = state.pathwayName.trim();
    const org = state.pathwayOrg.trim();
    const details = [
      ['Generated by', (name || 'Name not stated') + ', ' + (org || 'organisation not stated')],
      ['Date generated', dateText],
      ['Programmes', ids.length + ' programme' + (ids.length === 1 ? '' : 's')],
    ];
    details.forEach(function (d) {
      setText(PDF_COLOURS.inkFaint, 8, true);
      doc.text(d[0].toUpperCase(), M, y);
      const v = lines(d[1], CW - 120, 11, false);
      setText(PDF_COLOURS.ink, 11, false);
      doc.text(v, M + 120, y);
      y += Math.max(1, v.length) * 14.3 + 6;
    });
    y += 6;

    // ---- Notice ----
    const notice = 'This file is generated by the Service Training Roadmap Tool developed by Civil Service College. ' +
      'The information is accurate as at the date generated, but programme details, schedules, and availability can change. ' +
      'It is your responsibility to check that the information is current and accurate with the respective IDOs before you sign up for any programme.';
    const nLines = lines(notice, CW - 40, 9.5, false);
    const nH = 30 + nLines.length * 12.4;
    doc.setFillColor.apply(doc, PDF_COLOURS.panel);
    doc.roundedRect(M, y, CW, nH, 6, 6, 'F');
    setText(PDF_COLOURS.ink, 9.5, true);
    doc.text('Please note', M + 20, y + 20);
    setText(PDF_COLOURS.inkSoft, 9.5, false);
    doc.text(nLines, M + 20, y + 35);
    y += nH + 24;

    // ---- Section heading ----
    setText(PDF_COLOURS.ink, 14, true);
    doc.text('Your programmes', M, y);
    y += 8;
    doc.setDrawColor.apply(doc, PDF_COLOURS.border); doc.setLineWidth(1);
    doc.line(M, y, W - M, y);
    y += 18;

    // ---- Programme cards ----
    ids.forEach(function (id, i) {
      const p = progById[id];
      const textX = M + 44, textW = CW - 60;
      const titleL = lines(p.title, textW, 12.5, true);
      const meta = [levelCode(p.levels) + ' ' + p.levels.map(levelLabel).join(', '), p.modality, p.duration, p.id].join('  •  ');
      const metaL = lines(meta, textW, 9, false);
      const descL = lines(p.overview ? shortText(p.overview, 340) : NO_DATA_MESSAGE, textW, 9.5, false);
      const compL = lines('Competencies: ' + compNames(p.competencyIds), textW, 9, false);
      const roleL = lines('Recommended for: ' + roleNames(p.roleIds), textW, 9, false);
      const linkText = p.url ? p.url : (p.remarks || 'By nomination only') + '. See LEARN for details.';
      const linkL = lines(linkText, textW, 8.5, false);
      const h = 16 + titleL.length * 16.3 + metaL.length * 11.5 + 8 + descL.length * 12.4 + 8 + compL.length * 11.5 + roleL.length * 11.5 + 6 + linkL.length * 11 + 14;
      ensure(h + 10);

      // card
      doc.setDrawColor.apply(doc, PDF_COLOURS.border); doc.setLineWidth(0.8);
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(M, y, CW, h, 6, 6, 'FD');
      // left accent strip in the programme's first competency colour
      const accent = compById[p.competencyIds[0]] ? hexToRgb(compById[p.competencyIds[0]].colour) : PDF_COLOURS.bar;
      doc.setFillColor(accent[0], accent[1], accent[2]);
      doc.rect(M, y + 6, 3, h - 12, 'F');
      // number circle
      doc.setFillColor.apply(doc, PDF_COLOURS.bar);
      doc.circle(M + 24, y + 24, 11, 'F');
      setText([255, 255, 255], 10.5, true);
      doc.text(String(i + 1), M + 24, y + 27.5, { align: 'center' });

      let cy = y + 16 + 11;
      setText(PDF_COLOURS.ink, 12.5, true);
      doc.text(titleL, textX, cy); cy += titleL.length * 16.3 - 3;
      setText(PDF_COLOURS.inkFaint, 9, false);
      doc.text(metaL, textX, cy); cy += metaL.length * 11.5 + 8;
      setText(PDF_COLOURS.inkSoft, 9.5, false);
      doc.text(descL, textX, cy); cy += descL.length * 12.4 + 8;
      setText(PDF_COLOURS.ink, 9, false);
      doc.text(compL, textX, cy); cy += compL.length * 11.5;
      doc.text(roleL, textX, cy); cy += roleL.length * 11.5 + 6;
      if (p.url) {
        setText(PDF_COLOURS.link, 8.5, false);
        linkL.forEach(function (ln, k) { doc.textWithLink(ln, textX, cy + k * 11, { url: p.url }); });
      } else {
        setText(PDF_COLOURS.inkFaint, 8.5, false);
        doc.text(linkL, textX, cy);
      }
      y += h + 10;
    });

    // ---- Footer on every page ----
    const year = now.getFullYear();
    const pages = doc.getNumberOfPages();
    for (let pg = 1; pg <= pages; pg++) {
      doc.setPage(pg);
      doc.setDrawColor.apply(doc, PDF_COLOURS.border); doc.setLineWidth(0.8);
      doc.line(M, H - 34, W - M, H - 34);
      setText(PDF_COLOURS.inkFaint, 8.5, false);
      doc.text('© ' + year + ' Civil Service College. All Rights Reserved.', M, H - 20);
      doc.text('Generated ' + dateText + '  •  Page ' + pg + ' of ' + pages, W - M, H - 20, { align: 'right' });
    }

    const filename = 'service-training-pathway.pdf';
    let downloads = null;
    if (window.claude && typeof window.claude.use === 'function') {
      try { downloads = await window.claude.use('downloads'); } catch (e) { downloads = null; }
    }
    if (downloads) {
      try {
        await downloads.save({ filename: filename, data: doc.output('blob') });
      } catch (e) {
        if (!e || e.code !== 'declined') showToast('Could not save the PDF. Please try again.');
      }
    } else {
      doc.save(filename);
    }
  }

  /* ======================================================================
     THEME
  ====================================================================== */
  const themeSwitch = document.getElementById('theme-switch');
  const themeLabel = document.getElementById('theme-label');
  themeSwitch.addEventListener('click', function () {
    state.theme = state.theme === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', state.theme);
    themeSwitch.setAttribute('aria-checked', String(state.theme === 'dark'));
    themeLabel.textContent = state.theme === 'dark' ? 'Dark mode' : 'Light mode';
  });

  /* ======================================================================
     PROGRAMME DETAIL MODAL
  ====================================================================== */
  const modalOverlay = document.getElementById('modal-overlay');
  const modalTitle = document.getElementById('modal-title');
  const modalKicker = document.getElementById('modal-kicker');
  const modalBody = document.getElementById('modal-body');
  let currentModalProgId = null;
  document.querySelector('[data-close-modal]').addEventListener('click', closeModal);
  modalOverlay.addEventListener('click', function (e) { if (e.target === modalOverlay) closeModal(); });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (modalOverlay.classList.contains('open')) closeModal();
    if (pathwayOverlay.classList.contains('open')) pathwayOverlay.classList.remove('open');
    if (presetOverlay.classList.contains('open')) presetOverlay.classList.remove('open');
  });

  function refreshModalAddButton() {
    const holder = document.getElementById('modal-add-holder');
    if (holder) holder.innerHTML = addBtnHtml(currentModalProgId, '');
  }

  function openModal(programme) {
    currentModalProgId = programme.id;
    modalTitle.textContent = programme.title;
    modalKicker.innerHTML = programme.url ? '' : '<span class="pill-nomination">' + (programme.remarks || 'By nomination only') + '</span>';

    const metaBits = [programme.modality, programme.duration];
    const metaLine = '<p class="modal-meta">' + metaBits.join(' &middot; ') +
      (programme.status === 'Revamping' ? ' <span class="badge-revamp">Being revamped</span>' : '') + '</p>';

    const compTags = programme.competencyIds.map(function (id) {
      const c = compById[id];
      return '<span class="tag"><span class="dot" style="background:' + c.colour + '"></span>' + c.name + '</span>';
    }).join('');
    const valueTags = programme.valueIds.filter(function (id) { return valueById[id]; }).map(function (id) { return '<span class="tag">' + valueById[id].symbol + ' ' + escapeHtml(valueById[id].name) + '</span>'; }).join('');

    let extra = '';
    if (programme.criticalRaw || programme.relevantRaw) {
      extra = '<div class="modal-section"><h4 class="modal-section-head">Related competency framework</h4><div class="modal-extra">' +
        (programme.criticalRaw ? '<p><strong>Critical competencies mapped:</strong> ' + programme.criticalRaw + '</p>' : '') +
        (programme.relevantRaw ? '<p><strong>Relevant competencies mapped:</strong> ' + programme.relevantRaw + '</p>' : '') +
      '</div></div>';
    }

    // Prefer hand-written overview copy once it exists for a programme;
    // otherwise say plainly that there's nothing to show yet, rather than
    // generating a stand-in summary.
    const overviewHtml = programme.overview
      ? renderMultiParagraph(programme.overview)
      : '<p class="modal-overview">' + NO_DATA_MESSAGE + '</p>';

    const outcomesHtml =
      '<div class="modal-section"><h4 class="modal-section-head">Learning outcomes</h4>' +
      ((programme.learningOutcomes && programme.learningOutcomes.length > 1)
        ? '<ol class="modal-outcomes">' + programme.learningOutcomes.map(function (o) { return '<li>' + escapeHtml(o) + '</li>'; }).join('') + '</ol>'
        : (programme.learningOutcomes && programme.learningOutcomes.length === 1)
        ? '<p class="modal-overview">' + escapeHtml(programme.learningOutcomes[0]) + '</p>'
        : '<p class="modal-overview">' + NO_DATA_MESSAGE + '</p>') +
      '</div>';

    const outlineHtml =
      '<div class="modal-section"><details class="ref-section"><summary>Programme outline</summary><div class="ref-body modal-outline">' +
      (programme.outline ? renderOutline(programme.outline) : '<p>' + NO_DATA_MESSAGE + '</p>') +
      '</div></details></div>';

    modalBody.innerHTML =
      metaLine +
      '<div class="btn-row" style="margin-top:6px;">' +
      (programme.url
        ? '<a class="btn btn-primary" href="' + programme.url + '" target="_blank" rel="noopener">Open on LEARN</a>'
        : '<a class="btn btn-secondary" href="' + LEARN_HOME_URL + '" target="_blank" rel="noopener">Visit LEARN</a>') +
      '<span id="modal-add-holder">' + addBtnHtml(programme.id, '') + '</span>' +
      '</div>' +
      '<div class="modal-section"><h4 class="modal-section-head">Overview</h4>' + overviewHtml + '</div>' +
      outcomesHtml +
      '<div class="modal-section"><h4 class="modal-section-head">Details</h4>' +
        '<div class="modal-field-group"><div class="d-label">Competencies mapped</div><div class="tag-row" style="margin-top:6px;">' + compTags + '</div></div>' +
        '<div class="modal-field-group"><div class="d-label">Proficiency</div><div class="tag-row" style="margin-top:6px;"><span class="tag">' + levelCode(programme.levels) + ' \u2014 ' + programme.levels.map(levelLabel).join(', ') + '</span></div></div>' +
        '<div class="modal-field-group"><div class="d-label">Public service values</div><div class="tag-row" style="margin-top:6px;">' + valueTags + '</div></div>' +
        '<div class="detail-grid">' +
          '<div><div class="d-label">Modality</div><div class="d-value">' + programme.modality + '</div></div>' +
          '<div><div class="d-label">Duration</div><div class="d-value">' + programme.duration + '</div></div>' +
          '<div><div class="d-label">Recommended for</div><div class="d-value">' + roleNames(programme.roleIds) + '</div></div>' +
        '</div>' +
      '</div>' +
      outlineHtml +
      extra;
    modalOverlay.classList.add('open');
    document.querySelector('[data-close-modal]').focus();
  }
  function closeModal() { modalOverlay.classList.remove('open'); currentModalProgId = null; }

  /* Expandable result cards used inside preset pathways */
  function wireResultCardToggles(container) {
    container.querySelectorAll('[data-toggle-card]').forEach(function (b) {
      b.addEventListener('click', function () { document.getElementById(b.dataset.toggleCard).classList.toggle('expanded'); });
    });
    container.querySelectorAll('[data-open-detail]').forEach(function (b) {
      b.addEventListener('click', function () {
        // Close whichever modal this button lives in first — both the
        // preset and pathway overlays use the same full-screen backdrop
        // style as the programme detail modal, so having two open at once
        // would show two overlapping backdrops rather than a clean
        // "drill into details" step.
        presetOverlay.classList.remove('open');
        pathwayOverlay.classList.remove('open');
        openModal(progById[b.dataset.openDetail]);
      });
    });
  }

  /* ======================================================================
     PRESET PATHWAYS
     Each preset is a fixed, hand-picked list of real programmes, grouped
     into three categories (grade band, modality, or skill) — not a live
     query and not editable, since the point is to hand off something
     concrete a learner or their manager can act on and track against a
     LEARN playlist, not a starting point for a further conversation.
  ====================================================================== */
  function getPresetSteps(preset) {
    if (preset.category === 'band') {
      return programmes.filter(function (p) { return p.roleIds.indexOf(preset.role) !== -1; })
        .slice().sort(function (a, b) { return (Math.min.apply(null, a.levels) - Math.min.apply(null, b.levels)) || a.title.localeCompare(b.title); })
        .map(function (p) { return { programmeId: p.id, note: buildSkillNote(p) }; });
    }
    if (preset.category === 'skill') {
      return programmesForSkill(preset.skillAreaId).map(function (p) {
        return { programmeId: p.id, note: buildSkillNote(p) };
      });
    }
    return preset.steps;
  }

  function presetCardHtml(preset) {
    const stepCount = getPresetSteps(preset).length;
    // The whole card opens the pathway; the button gives it a clear label and keyboard focus.
    return '<div class="preset-card" data-preset="' + preset.id + '">' +
      '<div class="title">' + escapeHtml(preset.title) + '</div>' +
      (preset.subtitle ? '<p class="preset-sub">' + escapeHtml(preset.subtitle) + (preset.level ? ' &middot; L' + preset.level : '') + '</p>' : '') +
      '<p class="desc">' + escapeHtml(preset.desc) + '</p>' +
      '<p class="desc preset-count"><strong>' + stepCount + ' programme' + (stepCount === 1 ? '' : 's') + '</strong></p>' +
      '<button class="btn btn-secondary btn-sm preset-cta" aria-label="View the recommended programmes for ' + escapeHtml(preset.title) + '">View the recommended programmes <span aria-hidden="true">&rarr;</span></button>' +
    '</div>';
  }
  function renderPresets() {
    const groups = [
      { key: 'band', label: 'By role band', lead: 'Programmes recommended for your designation, from junior officers to directors.' },
      { key: 'skill', label: 'By skill', lead: 'Programmes that build one specific skill, across all levels.' },
    ];
    const filter = state.presetGroup || 'all';
    document.getElementById('preset-filter').innerHTML = [{ key: 'all', label: 'All' }].concat(groups).map(function (g) {
      return '<button class="browse-pill" aria-pressed="' + (filter === g.key) + '" data-preset-group="' + g.key + '">' + g.label + '</button>';
    }).join('');
    document.querySelectorAll('[data-preset-group]').forEach(function (b) {
      b.addEventListener('click', function () { state.presetGroup = b.dataset.presetGroup; renderPresets(); });
    });
    const holder = document.getElementById('preset-groups');
    holder.innerHTML = groups.filter(function (g) { return filter === 'all' || filter === g.key; }).map(function (g) {
      const items = presetPathways.filter(function (p) { return p.category === g.key; });
      if (!items.length) return '';
      return '<div class="preset-group"><h3 class="preset-group-head">' + g.label + ' <span class="preset-group-count">' + items.length + '</span></h3>' +
        '<p class="preset-group-lead">' + g.lead + '</p>' +
        '<div class="preset-grid">' + items.map(presetCardHtml).join('') + '</div></div>';
    }).join('') || '<p class="empty-state">No presets yet.</p>';
    holder.querySelectorAll('[data-preset]').forEach(function (b) {
      b.addEventListener('click', function () {
        state.selectedPreset = b.dataset.preset;
        renderPresetResults();
        presetOverlay.classList.add('open');
      });
    });
  }

  function presetStepCardHtml(presetId, stepIndex, programme, note) {
    const uid = 'preset-' + presetId + '-' + stepIndex + '-' + programme.id;
    return (
      '<div class="prog-card" id="' + uid + '">' +
        '<button class="prog-card-head" data-toggle-card="' + uid + '">' +
          '<span class="prog-rank">' + (stepIndex + 1) + '</span>' +
          '<span class="prog-main"><span class="top-line">' +
            '<span class="prog-title">' + programme.title + '</span>' +
            '<span class="tag">' + levelCode(programme.levels) + '</span>' +
          '</span><span class="prog-meta">' + programme.modality + ' &middot; ' + programme.duration + '</span></span>' +
          '<span class="chevron">\u25BC</span>' +
        '</button>' +
        '<div class="prog-card-body">' +
          '<p class="rationale"><strong>About this programme:</strong> ' + (programme.overview ? escapeHtml(programme.overview) : NO_DATA_MESSAGE) + '</p>' +
          '<p class="rationale" style="margin-top:8px;"><strong>Why it is relevant:</strong> ' + note + '</p>' +
          '<div class="detail-grid">' +
            '<div><div class="d-label">Competencies mapped</div><div class="d-value">' + compNames(programme.competencyIds) + '</div></div>' +
            '<div><div class="d-label">Recommended for</div><div class="d-value">' + roleNames(programme.roleIds) + '</div></div>' +
          '</div>' +
          '<div class="btn-row">' +
          (programme.url
            ? '<a class="btn btn-primary" href="' + programme.url + '" target="_blank" rel="noopener">Open on LEARN</a>'
            : '<a class="btn btn-secondary" href="' + LEARN_HOME_URL + '" target="_blank" rel="noopener">Visit LEARN</a>') +
          addBtnHtml(programme.id, '', note) +
          '<button class="btn btn-secondary" data-open-detail="' + programme.id + '">More details</button>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function unavailableStepHtml(stepIndex, title, reason) {
    return (
      '<div class="prog-card">' +
        '<div class="prog-card-head" style="cursor:default;">' +
          '<span class="prog-rank">' + (stepIndex + 1) + '</span>' +
          '<span class="prog-main"><span class="top-line">' +
            '<span class="badge-revamp">No longer available</span>' +
            '<span class="prog-title" style="color:var(--ink-faint);">' + escapeHtml(title) + '</span>' +
          '</span><span class="prog-meta">' + escapeHtml(reason) + '</span></span>' +
        '</div>' +
      '</div>'
    );
  }

  function renderPresetResults() {
    const preset = presetPathways.find(function (p) { return p.id === state.selectedPreset; });
    if (!preset) return;
    const steps = getPresetSteps(preset);

    let unavailableCount = 0;
    const cardsHtml = steps.map(function (s, i) {
      if (s.unavailableTitle) {
        unavailableCount++;
        return unavailableStepHtml(i, s.unavailableTitle, s.reason || 'No longer active in the current catalogue.');
      }
      const programme = progById[s.programmeId];
      if (!programme) {
        unavailableCount++;
        return unavailableStepHtml(i, s.programmeId, 'No longer active in the current catalogue.');
      }
      return presetStepCardHtml(preset.id, i, programme, s.note);
    }).join('');

    const allIds = steps.filter(function (s) { return s.programmeId && progById[s.programmeId]; }).map(function (s) { return s.programmeId; });
    const allAdded = allIds.length > 0 && allIds.every(function (id) { return state.pathway.has(id); });

    const unavailableNote = unavailableCount
      ? '<p class="chip-note" style="margin-top:16px; max-width:640px;">' + unavailableCount + ' programme' + (unavailableCount === 1 ? '' : 's') + ' from the original list ' + (unavailableCount === 1 ? 'is' : 'are') + ' no longer active in the current catalogue, and ' + (unavailableCount === 1 ? 'is' : 'are') + ' shown above for reference only, with no link or add-to-pathway option.</p>'
      : '';

    document.getElementById('preset-modal-title').textContent = preset.title;
    document.getElementById('preset-modal-body').innerHTML =
      '<div class="summary-strip">' +
        (preset.role ? '<div class="item"><span class="k">Role</span><span class="v">' + roleById[preset.role].name + '</span></div>' : '') +
        '<div class="item"><span class="k">Modality</span><span class="v">' + preset.modality + '</span></div>' +
        '<div class="item"><span class="k">Focus</span><span class="v">' + preset.skillFocus + '</span></div>' +
      '</div>' +
      '<div class="prog-list">' + cardsHtml + '</div>' +
      '<div class="btn-row">' +
        (allIds.length ? '<button class="btn btn-secondary btn-sm" id="preset-add-all"' + (allAdded ? ' disabled' : '') + '>' + (allAdded ? '\u2713 All added to pathway' : 'Add all ' + allIds.length + ' to pathway') + '</button>' : '') +
        '<button class="btn btn-secondary btn-sm" id="preset-open-pathway">Open my pathway</button>' +
        (preset.playlistUrl
          ? '<a class="btn btn-secondary btn-sm" href="' + preset.playlistUrl + '" target="_blank" rel="noopener">View LEARN playlist</a>'
          : '<button class="btn btn-secondary btn-sm" disabled title="No LEARN playlist linked yet">View LEARN playlist</button>') +
      '</div>' +
      (preset.playlistUrl
        ? ''
        : '<p class="chip-note" style="margin-top:16px; max-width:640px;">No LEARN playlist has been linked to this pathway yet. Please check with the programme owner.</p>') +
      unavailableNote;

    const body = document.getElementById('preset-modal-body');
    wireResultCardToggles(body);
    const addAllBtn = document.getElementById('preset-add-all');
    if (addAllBtn && !allAdded) {
      addAllBtn.addEventListener('click', function () {
        allIds.forEach(function (id) { if (!state.pathway.has(id)) state.pathway.set(id, null); });
        syncAfterPathwayChange();
      });
    }
    document.getElementById('preset-open-pathway').addEventListener('click', openPathway);
  }

  /* ======================================================================
     COMPETENCY & LEVEL MATRIX — a reference table for HR/programme owners,
     not part of the learner-facing flow. Every active programme against
     every competency, generated live from the same catalogue as
     everything else, so it can never drift out of sync with it.
  ====================================================================== */
  function renderMatrix() {
    // Filter dropdowns are only (re)populated once — their options don't
    // change between renders, only which one is selected does.
    const compSelect = document.getElementById('matrix-filter-comp');
    if (compSelect.options.length === 1) {
      compSelect.insertAdjacentHTML('beforeend', competencies.map(function (c) {
        return '<option value="' + c.id + '">' + c.name + '</option>';
      }).join(''));
    }
    const levelSelect = document.getElementById('matrix-filter-level');
    if (levelSelect.options.length === 1) {
      levelSelect.insertAdjacentHTML('beforeend', proficiencyLevels.map(function (l) {
        return '<option value="' + l.id + '">' + levelLabel(l.id) + '</option>';
      }).join(''));
    }
    const modalitySelect = document.getElementById('matrix-filter-modality');
    if (modalitySelect.options.length === 1) {
      const modalities = Array.from(new Set(programmes.map(function (p) { return p.modality; }))).sort();
      modalitySelect.insertAdjacentHTML('beforeend', modalities.map(function (m) {
        return '<option value="' + m + '">' + m + '</option>';
      }).join(''));
    }
    const roleSelect = document.getElementById('matrix-filter-role');
    if (roleSelect.options.length === 1) {
      roleSelect.insertAdjacentHTML('beforeend', roleBands.map(function (r) {
        return '<option value="' + r.id + '">' + r.name + '</option>';
      }).join(''));
    }

    const f = state.matrixFilters;
    document.getElementById('matrix-search').value = f.search;
    compSelect.value = f.comp;
    levelSelect.value = f.level;
    modalitySelect.value = f.modality;
    roleSelect.value = f.role;

    const filtered = programmes.filter(function (p) {
      if (f.search && p.title.toLowerCase().indexOf(f.search.toLowerCase()) === -1) return false;
      if (f.comp && !p.competencyIds.includes(f.comp)) return false;
      if (f.level && !p.levels.includes(Number(f.level))) return false;
      if (f.modality && p.modality !== f.modality) return false;
      if (f.role && !p.roleIds.includes(f.role)) return false;
      return true;
    }).sort(function (a, b) { return a.title.localeCompare(b.title); });

    document.getElementById('matrix-count').textContent = filtered.length + ' of ' + programmes.length + ' active programmes';

    const table = document.getElementById('matrix-table');
    const headHtml =
      '<thead><tr>' +
        '<th class="mx-prog-col">Programme</th>' +
        '<th class="mx-mod-col">Modality</th>' +
        competencies.map(function (c) {
          return '<th class="mx-comp-col"><span class="dot" style="background:' + c.colour + '"></span>' + escapeHtml(c.name) + '</th>';
        }).join('') +
      '</tr></thead>';

    const bodyHtml = '<tbody>' + (filtered.length
      ? filtered.map(function (p) {
          const rowLevel = levelCode(p.levels);
          return '<tr>' +
            '<td class="mx-prog-col"><div class="mx-prog-title">' + p.title + '</div><div class="mx-prog-meta">' + p.id + '</div></td>' +
            '<td class="mx-mod-col">' + p.modality + '</td>' +
            competencies.map(function (c) {
              return '<td class="mx-comp-col">' + (p.competencyIds.includes(c.id) ? levelCode((p.compLevels && p.compLevels[c.id]) || p.levels) : '') + '</td>';
            }).join('') +
          '</tr>';
        }).join('')
      : '<tr><td class="mx-prog-col" colspan="' + (2 + competencies.length) + '" style="text-align:center; color:var(--ink-faint);">No programmes match these filters.</td></tr>'
    ) + '</tbody>';

    table.innerHTML = headHtml + bodyHtml;

    document.getElementById('matrix-search').oninput = function (e) { state.matrixFilters.search = e.target.value; renderMatrix(); };
    compSelect.onchange = function (e) { state.matrixFilters.comp = e.target.value; renderMatrix(); };
    levelSelect.onchange = function (e) { state.matrixFilters.level = e.target.value; renderMatrix(); };
    modalitySelect.onchange = function (e) { state.matrixFilters.modality = e.target.value; renderMatrix(); };
    roleSelect.onchange = function (e) { state.matrixFilters.role = e.target.value; renderMatrix(); };
    document.getElementById('matrix-reset').onclick = function () {
      state.matrixFilters = { search: '', comp: '', level: '', modality: '', role: '' };
      renderMatrix();
    };
  }

  /* ======================================================================
     BROWSE VIEW — replaces the old separate "Filter & explore" and
     "Full roadmap" pages, which did the same job two different ways.
     One sidebar of filters, one search box, one sortable grid.
  ====================================================================== */
  // Quick-jump pills for competency, so learners can focus on one category
  // instead of scrolling the whole grouped catalogue. Single-select: picking
  // a pill replaces whatever competency filter was set, in the sidebar too,
  // since both read and write the same state.
  function renderBrowsePills() {
    const holder = document.getElementById('browse-pills');
    const selected = state.browseFilters.comp;
    const allActive = selected.size === 0;
    let html = '<button class="browse-pill" aria-pressed="' + allActive + '" data-pill="">All</button>';
    html += competencies.map(function (c) {
      const active = selected.has(c.id);
      return '<button class="browse-pill" aria-pressed="' + active + '" data-pill="' + c.id + '"><span class="dot" style="background:' + c.colour + '"></span>' + c.name + '</button>';
    }).join('');
    holder.innerHTML = html;
    holder.querySelectorAll('[data-pill]').forEach(function (b) {
      b.addEventListener('click', function () {
        if (b.dataset.pill) state.browseFilters.comp = new Set([b.dataset.pill]);
        else state.browseFilters.comp = new Set();
        renderBrowseSidebar();
        renderBrowsePills();
        renderBrowseResults();
      });
    });
  }

  function renderBrowseSidebar() {
    const sidebar = document.getElementById('browse-sidebar');
    function group(title, items, key, getId, getLabel) {
      const selected = state.browseFilters[key];
      const opts = items.map(function (item) {
        const id = getId(item);
        const checked = selected.has(id) ? 'checked' : '';
        return '<label class="filter-option"><input type="checkbox" data-key="' + key + '" value="' + id + '" ' + checked + '> ' + getLabel(item) + '</label>';
      }).join('');
      const head = '<h4>' + title + (selected.size ? ' (' + selected.size + ')' : '') + '</h4>' +
        (selected.size ? '<button class="filter-clear" data-clear-group="' + key + '">Clear</button>' : '');
      return '<div class="filter-group"><div class="filter-group-head">' + head + '</div>' + opts + '</div>';
    }
    sidebar.innerHTML =
      group('Proficiency level', proficiencyLevels, 'level', function (l) { return l.id; }, function (l) { return 'L' + l.id + ' ' + l.name; }) +
      group('Role band', roleBands, 'role', function (r) { return r.id; }, function (r) { return r.name; }) +
      group('Modality', modalityOptions, 'modality', function (m) { return m; }, function (m) { return m; }) +
      group('Public service value', serviceValues, 'value', function (v) { return v.id; }, function (v) { return v.symbol + ' ' + v.name; });

    sidebar.querySelectorAll('input[type="checkbox"]').forEach(function (cb) {
      cb.addEventListener('change', function () {
        const key = cb.dataset.key;
        const val = (key === 'level') ? Number(cb.value) : cb.value;
        if (cb.checked) state.browseFilters[key].add(val); else state.browseFilters[key].delete(val);
        renderBrowseSidebar();
        renderBrowsePills();
        renderBrowseResults();
      });
    });
    sidebar.querySelectorAll('[data-clear-group]').forEach(function (b) {
      b.addEventListener('click', function () {
        state.browseFilters[b.dataset.clearGroup].clear();
        renderBrowseSidebar();
        renderBrowsePills();
        renderBrowseResults();
      });
    });
  }

  function renderBrowseResults() {
    const f = state.browseFilters;
    const q = state.browseSearch.trim().toLowerCase();
    let filtered = programmes.filter(function (p) {
      if (f.comp.size && !p.competencyIds.some(function (c) { return f.comp.has(c); })) return false;
      if (f.level.size && !p.levels.some(function (l) { return f.level.has(l); })) return false;
      if (f.role.size && !p.roleIds.some(function (r) { return f.role.has(r); })) return false;
      if (f.modality.size && !f.modality.has(p.modality)) return false;
      if (f.value.size && !p.valueIds.some(function (v) { return f.value.has(v); })) return false;
      if (q && p.title.toLowerCase().indexOf(q) === -1) return false;
      return true;
    });

    const anyActive = f.comp.size || f.level.size || f.role.size || f.modality.size || f.value.size || q;
    document.getElementById('browse-count').textContent = anyActive
      ? filtered.length + ' of ' + programmes.length + ' active programmes match'
      : programmes.length + ' active programmes';

    renderFilterSummary();
    const grid = document.getElementById('browse-grid');
    if (filtered.length === 0) {
      grid.innerHTML = '<p class="empty-state">No programmes match this combination. Try clearing a filter or the search box.</p>';
      return;
    }

    function cardHtml(p) {
      const added = state.pathway.has(p.id);
      const compDots = p.competencyIds.map(function (id) {
        return '<span class="dot" style="background:' + compById[id].colour + '"></span>';
      }).join('');
      const dotsSpan = '<span class="tag-dots" title="' + compNames(p.competencyIds) + '">' + compDots + '</span>';
      return '<div class="explore-card reveal">' +
        '<button class="card-add-btn' + (added ? ' added' : '') + '" data-toggle-pathway="' + p.id + '" aria-label="' + (added ? 'Remove from pathway' : 'Add to pathway') + '">' + (added ? '\u2713' : '+') + '</button>' +
        '<button class="explore-card-open" data-open="' + p.id + '">' +
          '<div class="top-line">' +
            '<span class="tag">' + levelCode(p.levels) + '</span>' +
            '<span class="tag tag-plain">' + p.modality + '</span>' +
            dotsSpan +
            (p.status === 'Revamping' ? '<span class="badge-revamp">Being revamped</span>' : '') +
          '</div>' +
          '<div class="title">' + escapeHtml(p.title) + '</div>' +
          '<div class="meta">' + compNames(p.competencyIds) + '</div>' +
        '</button>' +
      '</div>';
    }

    let bodyHtml;
    if (state.browseSort === 'level') {
      filtered = filtered.slice().sort(function (a, b) { return Math.min(...a.levels) - Math.min(...b.levels); });
      bodyHtml = filtered.map(cardHtml).join('');
    } else if (state.browseSort === 'az') {
      filtered = filtered.slice().sort(function (a, b) { return a.title.localeCompare(b.title); });
      bodyHtml = filtered.map(cardHtml).join('');
    } else {
      // Group by competency, in catalogue order, then by level within each group.
      // When one or more competencies are filtered, only show headers for
      // those — otherwise a programme tagged with a second, unselected
      // competency was pulling in a group the person never asked for.
      const groupComps = f.comp.size ? competencies.filter(function (c) { return f.comp.has(c.id); }) : competencies;
      bodyHtml = groupComps.map(function (c) {
        const inGroup = filtered.filter(function (p) { return p.competencyIds.includes(c.id); })
          .slice().sort(function (a, b) { return Math.min(...a.levels) - Math.min(...b.levels); });
        if (inGroup.length === 0) return '';
        return '<div class="browse-group-head reveal"><span class="dot" style="background:' + c.colour + '"></span>' + c.name + '</div>' +
          inGroup.map(cardHtml).join('');
      }).join('');
    }

    grid.innerHTML = bodyHtml;
    observeReveal(grid);
    grid.querySelectorAll('[data-open]').forEach(function (btn) { btn.addEventListener('click', function () { openModal(progById[btn.dataset.open]); }); });
  }


  /* ======================================================================
     SMALL NOTICE (replaces browser alerts, which artifacts block)
  ====================================================================== */
  let toastTimer = null;
  function showToast(msg) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 4200);
  }

  /* ======================================================================
     SCROLL REVEAL — catalogue cards fade in as they come into view
  ====================================================================== */
  const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const revealObserver = ('IntersectionObserver' in window && !reduceMotion)
    ? new IntersectionObserver(function (entries) {
        let delay = 0;
        entries.forEach(function (en) {
          if (!en.isIntersecting) return;
          const el = en.target;
          el.style.transitionDelay = Math.min(delay, 360) + 'ms';
          delay += 45;
          el.classList.add('in');
          revealObserver.unobserve(el);
        });
      }, { rootMargin: '0px 0px -40px 0px', threshold: 0.08 })
    : null;
  function observeReveal(root) {
    root.querySelectorAll('.reveal:not(.in)').forEach(function (el) {
      if (revealObserver) revealObserver.observe(el); else el.classList.add('in');
    });
  }

  /* ======================================================================
     NAV — views are linked with #presets, #how-to-use and #matrix, so a
     page can be bookmarked or shared directly.
  ====================================================================== */
  const VIEW_HASH = { home: 'home', presets: 'presets', guide: 'how-to-use', cases: 'case-studies', matrix: 'matrix' };
  const tabs = document.querySelectorAll('nav.tabs button');
  tabs.forEach(function (tab) { tab.addEventListener('click', function () { switchView(tab.dataset.view); }); });
  document.addEventListener('click', function (e) {
    const g = e.target.closest('[data-goto]');
    if (g) { e.preventDefault(); switchView(g.dataset.goto); }
    const s = e.target.closest('[data-scroll-catalogue]');
    if (s) { e.preventDefault(); goToCatalogue(); }
  });

  function viewFromHash() {
    const h = (location.hash || '').replace('#', '');
    if (h === 'catalogue') return 'home';
    for (const v in VIEW_HASH) { if (VIEW_HASH[v] && VIEW_HASH[v] === h) return v; }
    return 'home';
  }
  window.addEventListener('hashchange', function () {
    const v = viewFromHash();
    if (v !== state.view) switchView(v, { fromHash: true });
  });

  function switchView(view, opts) {
    opts = opts || {};
    state.view = view;
    tabs.forEach(function (t) { t.setAttribute('aria-selected', String(t.dataset.view === view || (view === 'matrix' && t.dataset.view === 'home') || (view === 'cases' && t.dataset.view === 'guide'))); });
    document.querySelectorAll('.view').forEach(function (v) { v.classList.remove('active'); });
    document.getElementById('view-' + view).classList.add('active');
    if (view === 'home') renderHome();
    if (view === 'presets') renderPresets();
    if (view === 'guide') renderGuide();
    if (view === 'matrix') renderMatrix();
    if (view === 'cases') renderCases();
    // Always set a named hash (never clear it), so the viewer's address
    // bar and this page never disagree about which view is showing.
    if (!opts.fromHash) setHash(VIEW_HASH[view]);
    if (!opts.keepScroll) window.scrollTo(0, 0);
  }

  function setHash(name) {
    try { if (location.hash !== '#' + name) location.hash = name; } catch (e) { /* the hash is a convenience only */ }
  }
  function goToCatalogue() {
    if (state.view !== 'home') switchView('home', { keepScroll: true, fromHash: true });
    setHash('catalogue');
    const el = document.getElementById('catalogue');
    requestAnimationFrame(function () { el.scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' }); });
  }

  /* ======================================================================
     HOME — hero with a call to action, and the full catalogue below it
  ====================================================================== */
  function renderHome() {
    const levelsCount = proficiencyLevels.length;
    document.getElementById('home-stats').innerHTML =
      '<span><strong>' + programmes.length + '</strong> active programmes</span>' +
      '<span><strong>' + competencies.length + '</strong> functional competencies</span>' +
      '<span><strong>' + levelsCount + '</strong> proficiency levels</span>' +
      '<span><strong>' + presetPathways.length + '</strong> preset pathways</span>';
    renderBrowse();
    observeReveal(document.getElementById('view-home'));
  }

  /* ======================================================================
     BROWSE (on the home page). Controls are wired once, in wireBrowseControls.
  ====================================================================== */
  function renderBrowse() {
    renderBrowseSidebar();
    renderBrowsePills();
    renderBrowseResults();
    document.getElementById('browse-search').value = state.browseSearch;
    document.getElementById('browse-sort').value = state.browseSort;
  }
  function resetBrowseFilters() {
    state.browseFilters = { comp: new Set(), level: new Set(), role: new Set(), modality: new Set(), value: new Set() };
    state.browseSearch = '';
  }
  function wireBrowseControls() {
    document.getElementById('browse-search').addEventListener('input', function (e) { state.browseSearch = e.target.value; renderBrowseResults(); });
    document.getElementById('browse-sort').addEventListener('change', function (e) { state.browseSort = e.target.value; renderBrowseResults(); });
    document.getElementById('browse-reset').addEventListener('click', function () {
      resetBrowseFilters();
      renderBrowse();
    });
    const toggle = document.getElementById('filters-toggle');
    toggle.addEventListener('click', function () {
      const sb = document.getElementById('browse-sidebar');
      const open = sb.classList.toggle('open');
      toggle.setAttribute('aria-expanded', String(open));
      toggle.textContent = open ? 'Hide filters' : 'Show filters';
    });
  }

  // A clear line above the grid naming every filter that is switched on.
  function renderFilterSummary() {
    const f = state.browseFilters;
    const bits = [];
    f.comp.forEach(function (id) { if (compById[id]) bits.push('<span class="tag"><span class="dot" style="background:' + compById[id].colour + '"></span>' + escapeHtml(compById[id].name) + '</span>'); });
    f.level.forEach(function (id) { bits.push('<span class="tag">L' + id + ' ' + escapeHtml(levelLabel(id)) + '</span>'); });
    f.role.forEach(function (id) { if (roleById[id]) bits.push('<span class="tag">' + escapeHtml(roleById[id].name) + '</span>'); });
    f.modality.forEach(function (m) { bits.push('<span class="tag">' + escapeHtml(m) + '</span>'); });
    f.value.forEach(function (id) { if (valueById[id]) bits.push('<span class="tag">' + escapeHtml(valueById[id].name) + '</span>'); });
    const el = document.getElementById('filter-summary');
    if (!bits.length) { el.innerHTML = ''; el.hidden = true; return; }
    el.hidden = false;
    el.innerHTML = '<span class="fs-label">Filtered by:</span> ' + bits.join(' ') + ' <button class="link-btn" id="filter-summary-clear">Clear filters</button>';
    document.getElementById('filter-summary-clear').addEventListener('click', function () { resetBrowseFilters(); renderBrowse(); });
  }

  // Used by the guide: jump to the catalogue with a filter already applied.
  function showCatalogueWith(filters, note) {
    resetBrowseFilters();
    if (filters.comp) state.browseFilters.comp = new Set(filters.comp);
    if (filters.level) state.browseFilters.level = new Set(filters.level);
    if (state.view === 'home') renderBrowse();
    goToCatalogue();
    if (note) showToast(note);
  }

  /* ======================================================================
     HOW TO USE — every part of this page is read from the Excel file
  ====================================================================== */
  const guideSel = new Set(); // selected service function ids
  function renderGuide() {
    const numbered = guideSteps.filter(function (s) { return /^\d+$/.test(String(s.n).trim()); });
    const tips = guideSteps.filter(function (s) { return !/^\d+$/.test(String(s.n).trim()); });
    document.getElementById('guide-steps').innerHTML = numbered.map(function (s) {
      return '<li class="g-step"><span class="g-step-n">' + escapeHtml(s.n) + '</span><p>' + escapeHtml(s.text) + '</p></li>';
    }).join('');
    document.getElementById('guide-tips').innerHTML = tips.map(function (s) {
      return '<div class="g-tip"><strong>' + escapeHtml(s.n) + ':</strong> ' + escapeHtml(s.text) + '</div>';
    }).join('');

    // Part 1: service functions and areas of work
    const fnGrid = document.getElementById('guide-functions');
    fnGrid.innerHTML = serviceFunctions.length ? serviceFunctions.map(function (f) {
      const on = guideSel.has(f.id);
      return '<button class="fn-card' + (on ? ' on' : '') + '" aria-pressed="' + on + '" data-fn="' + f.id + '">' +
        '<span class="fn-check" aria-hidden="true">' + (on ? '✓' : '') + '</span>' +
        '<span class="fn-name">' + escapeHtml(f.name) + '</span>' +
        (f.desc ? '<span class="fn-desc">' + escapeHtml(f.desc) + '</span>' : '') +
        (f.areas.length ? '<span class="fn-areas-head">Areas of work</span><ol class="fn-areas">' + f.areas.map(function (a) { return '<li>' + escapeHtml(a) + '</li>'; }).join('') + '</ol>' : '') +
      '</button>';
    }).join('') : '<p class="empty-state">No service functions found in the Excel file.</p>';
    fnGrid.querySelectorAll('[data-fn]').forEach(function (b) {
      b.addEventListener('click', function () {
        if (guideSel.has(b.dataset.fn)) guideSel.delete(b.dataset.fn); else guideSel.add(b.dataset.fn);
        renderGuide();
      });
    });

    // Part 2: competency x service function table
    const ruleByComp = {};
    competencyRules.forEach(function (r) { ruleByComp[r.compId] = r; });
    const rowComps = competencies.filter(function (c) { return ruleByComp[c.id]; });
    const needed = new Set();
    rowComps.forEach(function (c) {
      const r = ruleByComp[c.id];
      if (r.critical.concat(r.relevant).some(function (id) { return guideSel.has(id); })) needed.add(c.id);
    });
    const table = document.getElementById('guide-comp-table');
    table.innerHTML =
      '<thead><tr><th class="gc-comp">Functional competency</th>' + serviceFunctions.map(function (f) {
        return '<th class="gc-fn' + (guideSel.has(f.id) ? ' sel' : '') + '">' + escapeHtml(f.name) + '</th>';
      }).join('') + '</tr></thead><tbody>' +
      rowComps.map(function (c) {
        const r = ruleByComp[c.id];
        return '<tr class="' + (needed.has(c.id) ? 'need' : '') + '"><td class="gc-comp"><span class="dot" style="background:' + c.colour + '"></span>' + escapeHtml(c.name) + '</td>' +
          serviceFunctions.map(function (f) {
            const crit = r.critical.indexOf(f.id) !== -1, rel = r.relevant.indexOf(f.id) !== -1;
            const label = crit ? 'Critical' : rel ? 'Relevant' : '';
            return '<td class="gc-cell' + (guideSel.has(f.id) ? ' sel' : '') + '"' + (label ? ' title="' + label + '"' : '') + '>' +
              (crit ? '<span class="mark crit">✓*</span>' : rel ? '<span class="mark">✓</span>' : '') + '</td>';
          }).join('') + '</tr>';
      }).join('') + '</tbody>';

    const action = document.getElementById('guide-comp-action');
    if (guideSel.size === 0) {
      action.innerHTML = '<p class="chip-note">Select one or more service functions above to highlight the competencies you need.</p>';
    } else {
      const ids = Array.from(needed);
      const count = programmes.filter(function (p) { return p.competencyIds.some(function (c) { return needed.has(c); }); }).length;
      action.innerHTML =
        '<p class="g-need">You need <strong>' + ids.length + '</strong> competenc' + (ids.length === 1 ? 'y' : 'ies') + ' for the service function' + (guideSel.size === 1 ? '' : 's') + ' you picked: ' +
        ids.map(function (id) { return '<span class="tag"><span class="dot" style="background:' + compById[id].colour + '"></span>' + escapeHtml(compById[id].name) + '</span>'; }).join(' ') + '</p>' +
        '<div class="btn-row">' +
          (count ? '<button class="btn btn-primary" id="guide-show-comps">Show the ' + count + ' matching programme' + (count === 1 ? '' : 's') + '</button>' : '<span class="chip-note">There are no active programmes for these competencies yet.</span>') +
          '<button class="btn btn-secondary" id="guide-clear">Clear selection</button>' +
        '</div>';
      const showBtn = document.getElementById('guide-show-comps');
      if (showBtn) showBtn.addEventListener('click', function () {
        showCatalogueWith({ comp: ids }, 'Showing programmes for the competencies you picked. Use Reset filters to see everything.');
      });
      document.getElementById('guide-clear').addEventListener('click', function () { guideSel.clear(); renderGuide(); });
    }

    // Part 3: proficiency levels
    document.getElementById('guide-levels').innerHTML = proficiencyLevels.map(function (l) {
      const count = programmes.filter(function (p) { return p.levels.indexOf(l.id) !== -1; }).length;
      return '<div class="lv-row">' +
        '<div class="lv-name" style="--lv:' + l.colour + '"><span class="lv-num">' + l.id + '</span>' + escapeHtml(l.name) + '</div>' +
        '<div class="lv-cell"><span class="lv-k">Typical designation</span>' + escapeHtml(l.designation || '—') + '</div>' +
        '<div class="lv-cell"><span class="lv-k">Typical work areas</span>' + escapeHtml(l.workAreas || l.desc || '—') + '</div>' +
        '<div class="lv-act">' + (count ? '<button class="btn btn-secondary btn-sm" data-level-go="' + l.id + '">See ' + count + ' programme' + (count === 1 ? '' : 's') + '</button>' : '<span class="chip-note">None yet</span>') + '</div>' +
      '</div>';
    }).join('');
    document.querySelectorAll('[data-level-go]').forEach(function (b) {
      b.addEventListener('click', function () {
        const id = Number(b.dataset.levelGo);
        showCatalogueWith({ level: [id] }, 'Showing programmes at Level ' + id + ' (' + levelLabel(id) + '). Use Reset filters to see everything.');
      });
    });

  }

  // Preview a different Excel file in this browser only. Nothing is uploaded
  // or saved; reloading the page goes back to the published file.
  function wireDataUpload() {
    const input = document.getElementById('data-file');
    document.querySelectorAll('[data-pick-file]').forEach(function (b) { b.addEventListener('click', function () { input.click(); }); });
    input.addEventListener('change', async function () {
      const f = input.files && input.files[0];
      if (!f) return;
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
        const cat = buildCatalogue(wb);
        if (!cat.programmes.length) throw new Error('No active programmes were found in that file.');
        applyCatalogue(cat, f.name + ' (preview in this browser only)');
        resetBrowseFilters();
        state.matrixFilters = { search: '', comp: '', level: '', modality: '', role: '' };
        resetMatrixSelects();
        Array.from(state.pathway.keys()).forEach(function (id) { if (!progById[id]) state.pathway.delete(id); });
        pathwayCountEl.textContent = String(state.pathway.size);
        guideSel.clear();
        if (document.body.classList.contains('load-failed')) {
          document.body.classList.remove('load-failed');
          switchView('home');
        } else {
          switchView(state.view, { keepScroll: true });
        }
        showToast('Loaded ' + f.name + '. This preview is only in your browser.');
      } catch (e) {
        showToast('Could not read that file: ' + (e && e.message ? e.message : 'unknown error'));
      }
      input.value = '';
    });
  }
  function resetMatrixSelects() {
    ['matrix-filter-comp', 'matrix-filter-level', 'matrix-filter-modality', 'matrix-filter-role'].forEach(function (id) {
      const s = document.getElementById(id);
      while (s.options.length > 1) s.remove(1);
    });
  }

  /* ======================================================================
     CASE STUDIES — the stories and needs come from the "Guide - Case Studies"
     sheet; the matching programmes are worked out live from the catalogue.
     A programme matches a need when the Master List pegs it at that exact
     level for that competency. "Stretch" shows the next level up.
  ====================================================================== */
  function programmesAt(compId, level) {
    return programmes.filter(function (p) { return p.compLevels && p.compLevels[compId] && p.compLevels[compId].indexOf(level) !== -1; });
  }
  function caseTableHtml(cs, levelOf, exclude) {
    exclude = exclude || [];
    // levelOf(need) gives the level to look at for that competency, or null to skip it
    const needs = cs.needs.filter(function (n) { return levelOf(n) != null; });
    const levels = Array.from(new Set(needs.map(levelOf))).sort();
    if (!levels.length) return '';
    const cols = cs.needs;
    let total = 0;
    const sections = levels.map(function (L) {
      const active = needs.filter(function (n) { return levelOf(n) === L; });
      const rows = programmes.filter(function (p) {
        return exclude.indexOf(p.id) === -1 && active.some(function (n) { return programmesAt(n.compId, L).indexOf(p) !== -1; });
      }).sort(function (a, b) {
        const ca = active.filter(function (n) { return (a.compLevels[n.compId] || []).indexOf(L) !== -1; }).length;
        const cb = active.filter(function (n) { return (b.compLevels[n.compId] || []).indexOf(L) !== -1; }).length;
        return (cb - ca) || a.title.localeCompare(b.title);
      });
      total += rows.length;
      if (!rows.length) {
        if (exclude.length) return '';
        return '<tr class="ct-level-row"><th class="ct-level" scope="rowgroup">L' + L + ' ' + escapeHtml(levelLabel(L)) + '</th>' +
          '<td colspan="' + (cols.length + 2) + '" class="ct-empty">No active programme is pegged at this level for ' + escapeHtml(joinWithAnd(active.map(function (n) { return compById[n.compId].name; }))) + ' yet.</td></tr>';
      }
      return rows.map(function (p, i) {
        return '<tr>' +
          (i === 0 ? '<th class="ct-level" scope="rowgroup" rowspan="' + rows.length + '"><span>L' + L + ' ' + escapeHtml(levelLabel(L)) + '</span></th>' : '') +
          '<td class="ct-prog"><button class="ct-title" data-open="' + p.id + '">' + escapeHtml(p.title) + '</button>' +
            '<div class="ct-meta">' + escapeHtml(p.id) + ' &middot; ' + escapeHtml(p.modality) + ' &middot; ' + escapeHtml(p.duration) +
            (p.status === 'Revamping' ? ' <span class="badge-revamp">Being revamped</span>' : '') +
            (!p.url && p.remarks ? ' <span class="pill-nomination">' + escapeHtml(p.remarks) + '</span>' : '') + '</div></td>' +
          cols.map(function (n) {
            const hit = levelOf(n) === L && (p.compLevels[n.compId] || []).indexOf(L) !== -1;
            return '<td class="ct-cell' + (levelOf(n) === L ? '' : ' off') + '">' + (hit ? '<span class="mark" aria-label="Yes">✓</span>' : '') + '</td>';
          }).join('') +
          '<td class="ct-add">' + addBtnHtml(p.id, 'btn-sm') + '</td>' +
        '</tr>';
      }).join('');
    }).join('');
    return { total: total, html:
      '<div class="ct-scroll"><table class="ct-table"><thead><tr>' +
        '<th class="ct-level-h">Level</th><th class="ct-prog-h">Programme</th>' +
        cols.map(function (n) { return '<th>' + escapeHtml(compById[n.compId].name) + '</th>'; }).join('') +
        '<th class="ct-add-h"><span class="sr-only">Add</span></th>' +
      '</tr></thead><tbody>' + sections + '</tbody></table></div>' };
  }
  function joinWithAnd(items) { return items.length > 1 ? items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1] : (items[0] || ''); }

  function renderCases() {
    const picker = document.getElementById('case-picker');
    const body = document.getElementById('case-body');
    if (!caseStudies.length) {
      picker.innerHTML = '';
      body.innerHTML = '<p class="empty-state">No case studies found. Add them to the "Guide - Case Studies" sheet in the Excel file.</p>';
      return;
    }
    if (state.caseIdx == null || state.caseIdx >= caseStudies.length) state.caseIdx = 0;
    picker.innerHTML = caseStudies.map(function (c, i) {
      return '<button role="tab" class="case-tab" aria-selected="' + (i === state.caseIdx) + '" data-case="' + i + '">' +
        '<span class="case-tab-n">Case ' + escapeHtml(c.n) + '</span><span class="case-tab-t">' + escapeHtml(c.title) + '</span></button>';
    }).join('');
    picker.querySelectorAll('[data-case]').forEach(function (b) {
      b.addEventListener('click', function () { state.caseIdx = Number(b.dataset.case); renderCases(); });
    });

    const cs = caseStudies[state.caseIdx];
    const needByComp = {};
    cs.needs.forEach(function (n) { needByComp[n.compId] = n.level; });
    const fnTags = cs.functions.map(function (id) { return '<span class="tag">' + escapeHtml(fnById[id] ? fnById[id].name : id) + '</span>'; }).join(' ');
    const compTags = cs.needs.map(function (n) { const c = compById[n.compId]; return '<span class="tag"><span class="dot" style="background:' + c.colour + '"></span>' + escapeHtml(c.name) + '</span>'; }).join(' ');
    const levelRows = competencies.map(function (c, i) {
      const L = needByComp[c.id];
      return '<tr class="' + (L ? 'need' : '') + '"><td class="cl-n">' + (i + 1) + '</td><td>' + escapeHtml(c.name) + '</td><td class="cl-l">' + (L ? 'L' + L + ' ' + escapeHtml(levelLabel(L)) : '') + '</td></tr>';
    }).join('');
    const main = caseTableHtml(cs, function (n) { return n.level; });
    const maxLevel = proficiencyLevels.length ? proficiencyLevels[proficiencyLevels.length - 1].id : 5;
    const mainIds = [];
    cs.needs.forEach(function (n) { programmesAt(n.compId, n.level).forEach(function (p) { if (mainIds.indexOf(p.id) === -1) mainIds.push(p.id); }); });
    // Stretch shows only programmes that are not already recommended above.
    const stretch = caseTableHtml(cs, function (n) { return n.level < maxLevel ? n.level + 1 : null; }, mainIds);
    const allAdded = mainIds.length && mainIds.every(function (id) { return state.pathway.has(id); });

    body.innerHTML =
      '<article class="case">' +
        '<div class="case-story">' +
          '<div class="case-avatar" aria-hidden="true">' + escapeHtml((cs.persona || '?').charAt(0)) + '</div>' +
          '<div><div class="case-kicker">Case ' + escapeHtml(cs.n) + ' &middot; ' + escapeHtml(cs.designation || '') + '</div>' +
          '<h3 class="case-title">' + escapeHtml(cs.title.charAt(0).toUpperCase() + cs.title.slice(1)) + '</h3>' +
          '<p class="case-text">' + escapeHtml(cs.story) + '</p></div>' +
        '</div>' +
        '<ol class="case-q">' +
          '<li><h4>Which Service Function(s) match ' + escapeHtml(cs.persona || 'this officer') + '’s work?</h4><div class="tag-row">' + (fnTags || '<span class="chip-note">Not stated</span>') + '</div></li>' +
          '<li><h4>Which functional competencies does ' + escapeHtml(cs.persona || 'this officer') + ' want to build?</h4><div class="tag-row">' + compTags + '</div></li>' +
          '<li><h4>What proficiency level does ' + escapeHtml(cs.persona || 'this officer') + ' want to reach in each?</h4>' +
            '<div class="cl-wrap"><table class="cl-table"><thead><tr><th>S/N</th><th>Functional competency</th><th>Required level</th></tr></thead><tbody>' + levelRows + '</tbody></table></div></li>' +
          '<li><h4>Which programmes should ' + escapeHtml(cs.persona || 'this officer') + ' focus on?</h4>' +
            '<p class="case-note">These are the active programmes pegged at the required level for each competency. Discuss them with your reporting officer, and check LEARN for run dates.</p>' +
            (main && main.total ? main.html : '<p class="empty-state">No active programmes match these needs yet.</p>') +
            '<div class="btn-row" style="margin-top:12px;">' +
              (mainIds.length ? '<button class="btn btn-primary btn-sm" id="case-add-all"' + (allAdded ? ' disabled' : '') + '>' + (allAdded ? '✓ All ' + mainIds.length + ' in your pathway' : 'Add all ' + mainIds.length + ' to my pathway') + '</button>' : '') +
              '<button class="btn btn-secondary btn-sm" data-goto="guide">Try the steps for your own role</button>' +
            '</div>' +
            (stretch && stretch.total ? '<details class="ref-section case-stretch"><summary>Want to stretch? Programmes one level up (' + stretch.total + ')</summary><div class="ref-body">' +
              '<p class="case-note">These programmes are one level up and are not already listed above. You are welcome to attend a programme pegged at a higher level if you want to upskill.</p>' + stretch.html + '</div></details>' : '') +
          '</li>' +
        '</ol>' +
      '</article>';

    body.querySelectorAll('[data-open]').forEach(function (b) { b.addEventListener('click', function () { openModal(progById[b.dataset.open]); }); });
    const addAll = document.getElementById('case-add-all');
    if (addAll && !allAdded) addAll.addEventListener('click', function () {
      mainIds.forEach(function (id) { if (!state.pathway.has(id)) state.pathway.set(id, null); });
      syncAfterPathwayChange();
      showToast('Added ' + mainIds.length + ' programmes to your pathway.');
    });
  }

  /* ======================================================================
     START AGAIN — the logo and app name go home and clear every choice
     (filters, search, sort, guide picks, case and preset choices, open
     windows). Programmes already in My pathway are kept.
  ====================================================================== */
  function resetEverything() {
    resetBrowseFilters();
    state.browseSort = 'competency';
    state.matrixFilters = { search: '', comp: '', level: '', modality: '', role: '' };
    state.presetGroup = 'all';
    state.caseIdx = 0;
    state.selectedPreset = null;
    guideSel.clear();
    pathwayOpen.clear();
    closeModal();
    presetOverlay.classList.remove('open');
    pathwayOverlay.classList.remove('open');
    const sb = document.getElementById('browse-sidebar');
    const ft = document.getElementById('filters-toggle');
    if (sb && sb.classList.contains('open')) { sb.classList.remove('open'); ft.setAttribute('aria-expanded', 'false'); ft.textContent = 'Show filters'; }
    document.getElementById('toast').classList.remove('show');
  }
  document.getElementById('brand-home').addEventListener('click', function (e) {
    e.preventDefault();
    resetEverything();
    switchView('home');
    window.scrollTo({ top: 0, behavior: 'auto' });
  });

  /* ======================================================================
     INIT — load the Excel file, then draw the page
  ====================================================================== */
  wireBrowseControls();
  wireDataUpload();
  document.getElementById('sf-year').textContent = String(new Date().getFullYear());
  // Back-to-top button: appears once you have scrolled down a little.
  const toTop = document.getElementById('to-top');
  function syncToTop() { toTop.classList.toggle('show', window.scrollY > 500); }
  window.addEventListener('scroll', syncToTop, { passive: true });
  toTop.addEventListener('click', function () { window.scrollTo({ top: 0, behavior: reduceMotion ? 'auto' : 'smooth' }); });
  syncToTop();
  // Fonts are added here rather than in the page head, so a slow or blocked font service
  // can never hold up the page. Without them the site uses the computer's standard font.
  (function () {
    const l = document.createElement('link');
    l.rel = 'stylesheet';
    l.href = 'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600;700&display=swap';
    document.head.appendChild(l);
  })();

  // Load a script, giving up after a time limit, so a blocked or silent network never leaves the page waiting.
  function loadScript(src, timeoutMs) {
    return new Promise(function (resolve, reject) {
      const el = document.createElement('script');
      let done = false;
      const timer = setTimeout(function () { if (!done) { done = true; reject(new Error('timed out')); } }, timeoutMs || 8000);
      el.src = src;
      el.onload = function () { if (!done) { done = true; clearTimeout(timer); resolve(); } };
      el.onerror = function () { if (!done) { done = true; clearTimeout(timer); reject(new Error('blocked or not found')); } };
      document.head.appendChild(el);
    });
  }
  // Tries each source in turn until the library is available. A copy in the vendor folder is
  // tried first if there is one, then cdnjs, then jsDelivr.
  async function loadFirst(sources, isReady, label) {
    const tried = [];
    for (let i = 0; i < sources.length; i++) {
      if (isReady()) return;
      try { await loadScript(sources[i], i === 0 ? 4000 : 8000); } catch (e) { tried.push(sources[i] + ' (' + e.message + ')'); }
      if (isReady()) return;
    }
    if (!isReady()) throw new Error('Could not load ' + label + '. Tried: ' + tried.join('; ') + '. Your network or server may be blocking these addresses.');
  }
  const XLSX_SOURCES = ['vendor/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'];
  const JSPDF_SOURCES = ['vendor/jspdf.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/3.0.2/jspdf.umd.min.js', 'https://cdn.jsdelivr.net/npm/jspdf@3.0.2/dist/jspdf.umd.min.js'];
  function jspdfReady() { return !!(window.jspdf && window.jspdf.jsPDF); }

  // Some servers block style="" written into the page. Re-apply those styles through the
  // style object, which such policies allow, so colour dots and similar details still show.
  if ('MutationObserver' in window) {
    new MutationObserver(function (list) {
      list.forEach(function (m) {
        m.addedNodes.forEach(function (n) {
          if (n.nodeType !== 1) return;
          const els = [n].concat(Array.prototype.slice.call(n.querySelectorAll('[style]')));
          els.forEach(function (el) {
            const css = el.getAttribute && el.getAttribute('style');
            if (css && !el.style.length) { try { el.style.cssText = css; } catch (e) { /* ignore */ } }
          });
        });
      });
    }).observe(document.body, { childList: true, subtree: true });
  }

  try {
    await loadFirst(XLSX_SOURCES, function () { return typeof XLSX !== 'undefined'; }, 'the spreadsheet reader (SheetJS)');
    // The PDF tool is only needed for Export as PDF, so load it in the background.
    loadFirst(JSPDF_SOURCES, jspdfReady, 'the PDF tool (jsPDF)').catch(function () { /* retried on export */ });
    const wb = await fetchWorkbook();
    applyCatalogue(buildCatalogue(wb), dataOrigin);
    document.body.classList.remove('loading');
    switchView(viewFromHash(), { fromHash: true });
    if ((location.hash || '') === '#catalogue') goToCatalogue();
  } catch (err) {
    document.body.classList.remove('loading');
    document.body.classList.add('load-failed');
    document.getElementById('load-error-msg').textContent = err && err.message ? err.message : 'Unknown error.';
  }

})();
