export function json(data, init = {}) {
  return Response.json(data, {
    ...init,
    headers: {
      'cache-control': 'no-store',
      ...(init.headers ?? {}),
    },
  });
}

export function requireDb(env) {
  if (!env.DB) {
    throw new Error('Cloudflare D1 binding DB is not configured.');
  }
  return env.DB;
}

export async function readJson(request) {
  try {
    return await request.json();
  } catch {
    return null;
  }
}

export function parseData(row) {
  if (!row?.data) return null;
  try {
    return JSON.parse(row.data);
  } catch {
    return null;
  }
}

export function rowId(prefix = 'id') {
  return `${prefix}-${crypto.randomUUID()}`;
}

export async function insertProject(db, project) {
  const projectId = project.project_id ?? project.projectId ?? rowId('project');
  const projectName = project.projectName ?? project.project_name ?? '';
  const data = { ...project, project_id: projectId };
  await db
    .prepare(
      `INSERT OR REPLACE INTO projects (project_id, project_name, data, updated_at)
       VALUES (?, ?, ?, CURRENT_TIMESTAMP)`
    )
    .bind(projectId, projectName, JSON.stringify(data))
    .run();
  return data;
}

export async function insertAnalysisResult(db, record) {
  const analysisId = record.analysis_id ?? record.id ?? rowId('analysis');
  const data = { ...record, analysis_id: analysisId, id: record.id ?? analysisId };
  await db
    .prepare(
      `INSERT OR REPLACE INTO analysis_results
       (analysis_id, project_id, project_name, analysis_date, total_volume, total_carbon, data, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`
    )
    .bind(
      analysisId,
      data.project_id ?? '',
      data.project_name ?? data.projectName ?? '',
      data.analysis_date ?? data.analysisDate ?? '',
      Number(data.totalVolume ?? 0),
      Number(data.totalCarbon ?? 0),
      JSON.stringify(data)
    )
    .run();
  return data;
}
