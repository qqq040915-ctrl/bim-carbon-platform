import { insertProject, json, readJson, requireDb } from '../_shared.js';

export async function onRequestPut({ request, env, params }) {
  try {
    const db = requireDb(env);
    const body = await readJson(request);
    if (!body) return json({ ok: false, error: 'Invalid JSON body.' }, { status: 400 });
    const project = await insertProject(db, { ...body, project_id: params.id });
    return json({ ok: true, project });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}

export async function onRequestDelete({ env, params }) {
  try {
    const db = requireDb(env);
    await db.prepare('DELETE FROM projects WHERE project_id = ?').bind(params.id).run();
    await db.prepare('DELETE FROM bim_files WHERE project_id = ?').bind(params.id).run();
    await db.prepare('DELETE FROM material_mappings WHERE project_id = ?').bind(params.id).run();
    await db.prepare('DELETE FROM material_schemes WHERE project_id = ?').bind(params.id).run();
    await db.prepare('DELETE FROM analysis_results WHERE project_id = ?').bind(params.id).run();
    return json({ ok: true });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}
