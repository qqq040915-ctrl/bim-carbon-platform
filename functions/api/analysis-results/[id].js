import { json, requireDb } from '../_shared.js';

export async function onRequestDelete({ env, params }) {
  try {
    const db = requireDb(env);
    await db.prepare('DELETE FROM analysis_results WHERE analysis_id = ?').bind(params.id).run();
    return json({ ok: true });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}
