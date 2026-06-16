import { insertAnalysisResult, json, parseData, readJson, requireDb } from '../_shared.js';

export async function onRequestGet({ request, env }) {
  try {
    const db = requireDb(env);
    const url = new URL(request.url);
    const projectId = url.searchParams.get('project_id');
    const statement = projectId
      ? db.prepare('SELECT data FROM analysis_results WHERE project_id = ? ORDER BY updated_at DESC').bind(projectId)
      : db.prepare('SELECT data FROM analysis_results ORDER BY updated_at DESC');
    const { results } = await statement.all();
    return json({ ok: true, analysisResults: results.map(parseData).filter(Boolean) });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}

export async function onRequestPost({ request, env }) {
  try {
    const db = requireDb(env);
    const body = await readJson(request);
    if (!body) return json({ ok: false, error: 'Invalid JSON body.' }, { status: 400 });
    const analysisResult = await insertAnalysisResult(db, body);
    return json({ ok: true, analysisResult }, { status: 201 });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}
