import { insertProject, json, parseData, readJson, requireDb } from '../_shared.js';

export async function onRequestGet({ env }) {
  try {
    const db = requireDb(env);
    const { results } = await db.prepare('SELECT data FROM projects ORDER BY updated_at DESC').all();
    return json({ ok: true, projects: results.map(parseData).filter(Boolean) });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}

export async function onRequestPost({ request, env }) {
  try {
    const db = requireDb(env);
    const body = await readJson(request);
    if (!body) return json({ ok: false, error: 'Invalid JSON body.' }, { status: 400 });
    const project = await insertProject(db, body);
    return json({ ok: true, project }, { status: 201 });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}
