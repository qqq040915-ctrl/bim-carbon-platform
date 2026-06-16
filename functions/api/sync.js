import { insertAnalysisResult, insertProject, json, parseData, readJson, requireDb } from './_shared.js';

async function tableData(db, table, orderBy = 'updated_at DESC') {
  const { results } = await db.prepare(`SELECT data FROM ${table} ORDER BY ${orderBy}`).all();
  return results.map(parseData).filter(Boolean);
}

async function clearAndInsert(db, table, rows, insertSql, bindValues) {
  await db.prepare(`DELETE FROM ${table}`).run();
  for (const row of rows) {
    await db.prepare(insertSql).bind(...bindValues(row)).run();
  }
}

export async function onRequestGet({ env }) {
  try {
    const db = requireDb(env);
    const stateRow = await db.prepare("SELECT active_project_id, calculation, data FROM app_state WHERE id = 'global'").first();
    const state = stateRow?.data ? JSON.parse(stateRow.data) : {};
    return json({
      ok: true,
      projectDatabase: await tableData(db, 'projects'),
      activeProjectId: stateRow?.active_project_id ?? state.activeProjectId ?? '',
      bimImportSummary: await tableData(db, 'bim_files'),
      bimRows: state.bimRows ?? [],
      materialMappingDatabase: await tableData(db, 'material_mappings'),
      materialPlans: await tableData(db, 'material_schemes'),
      history: await tableData(db, 'analysis_results'),
      calculation: stateRow?.calculation ? JSON.parse(stateRow.calculation) : state.calculation ?? null,
    });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}

export async function onRequestPost({ request, env }) {
  try {
    const db = requireDb(env);
    const body = await readJson(request);
    if (!body) return json({ ok: false, error: 'Invalid JSON body.' }, { status: 400 });

    const projects = Array.isArray(body.projectDatabase) ? body.projectDatabase : [];
    const bimFiles = Array.isArray(body.bimImportSummary) ? body.bimImportSummary : [];
    const mappings = Array.isArray(body.materialMappingDatabase) ? body.materialMappingDatabase : [];
    const schemes = Array.isArray(body.materialPlans) ? body.materialPlans : [];
    const history = Array.isArray(body.history) ? body.history : [];

    await db.prepare('DELETE FROM projects').run();
    for (const project of projects) await insertProject(db, project);

    await clearAndInsert(
      db,
      'bim_files',
      bimFiles,
      `INSERT OR REPLACE INTO bim_files (id, project_id, data, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP)`,
      (row) => [row.id ?? crypto.randomUUID(), row.project_id ?? '', JSON.stringify(row)]
    );

    await clearAndInsert(
      db,
      'material_mappings',
      mappings,
      `INSERT OR REPLACE INTO material_mappings
       (id, project_id, bim_material_name, item_name, data, updated_at)
       VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
      (row) => [
        row.id ?? `${row.project_id ?? 'mapping'}-${row.bimMaterialName ?? crypto.randomUUID()}`,
        row.project_id ?? '',
        row.bimMaterialName ?? row.bim_material_name ?? '',
        row.itemName ?? row.item_name ?? '',
        JSON.stringify(row),
      ]
    );

    await clearAndInsert(
      db,
      'material_schemes',
      schemes,
      `INSERT OR REPLACE INTO material_schemes
       (id, project_id, scheme_name, data, updated_at)
       VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)`,
      (row) => [row.id ?? crypto.randomUUID(), row.project_id ?? '', row.name ?? row.scheme_name ?? '', JSON.stringify(row)]
    );

    await db.prepare('DELETE FROM analysis_results').run();
    for (const record of history) await insertAnalysisResult(db, record);

    await db
      .prepare(
        `INSERT OR REPLACE INTO app_state (id, active_project_id, calculation, data, updated_at)
         VALUES ('global', ?, ?, ?, CURRENT_TIMESTAMP)`
      )
      .bind(
        body.activeProjectId ?? body.activeProject?.project_id ?? '',
        body.calculation ? JSON.stringify(body.calculation) : null,
        JSON.stringify({ bimRows: body.bimRows ?? [], calculation: body.calculation ?? null })
      )
      .run();

    return json({ ok: true });
  } catch (error) {
    return json({ ok: false, error: error.message }, { status: 500 });
  }
}
