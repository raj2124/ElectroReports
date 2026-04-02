const fs = require('fs');
const path = require('path');

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function createReportStore({ filePath }) {
  const resolvedFilePath = filePath;
  ensureDir(path.dirname(resolvedFilePath));

  function readAll() {
    try {
      if (!fs.existsSync(resolvedFilePath)) {
        return [];
      }
      const raw = fs.readFileSync(resolvedFilePath, 'utf8');
      if (!raw.trim()) {
        return [];
      }
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (_error) {
      return [];
    }
  }

  function writeAll(records) {
    fs.writeFileSync(resolvedFilePath, JSON.stringify(records, null, 2), 'utf8');
  }

  function sortByRecent(records) {
    return [...records].sort((a, b) => {
      return Date.parse(String(b.updatedAt || b.createdAt || 0)) - Date.parse(String(a.updatedAt || a.createdAt || 0));
    });
  }

  function buildId() {
    const timestamp = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
    const suffix = Math.floor(1000 + Math.random() * 9000);
    return `ER-${timestamp}-${suffix}`;
  }

  function listReports(query = '') {
    const all = sortByRecent(readAll());
    const text = String(query || '').trim().toLowerCase();
    if (!text) {
      return all;
    }
    return all.filter((report) => {
      return [
        report?.project?.projectNo,
        report?.project?.clientName,
        report?.project?.siteLocation,
        report?.project?.workOrder,
        report?.project?.engineerName
      ]
        .join(' ')
        .toLowerCase()
        .includes(text);
    });
  }

  function getReport(id) {
    return readAll().find((report) => String(report.id) === String(id)) || null;
  }

  function addReport(reportInput) {
    const records = readAll();
    const now = new Date().toISOString();
    const report = {
      ...reportInput,
      id: buildId(),
      createdAt: now,
      updatedAt: now
    };
    writeAll(sortByRecent([report, ...records]));
    return report;
  }

  function updateReport(id, updater) {
    const records = readAll();
    const index = records.findIndex((report) => String(report.id) === String(id));
    if (index === -1) {
      return null;
    }
    const current = records[index];
    const nextValue =
      typeof updater === 'function'
        ? updater(current)
        : {
            ...current,
            ...(updater || {})
          };
    const updated = {
      ...current,
      ...(nextValue || {}),
      id: current.id,
      createdAt: current.createdAt,
      updatedAt: new Date().toISOString()
    };
    const nextRecords = [...records];
    nextRecords[index] = updated;
    writeAll(sortByRecent(nextRecords));
    return updated;
  }

  function deleteReport(id) {
    const records = readAll();
    const target = records.find((report) => String(report.id) === String(id)) || null;
    if (!target) {
      return { deleted: false, report: null };
    }
    const remaining = records.filter((report) => String(report.id) !== String(id));
    writeAll(sortByRecent(remaining));
    return { deleted: true, report: target };
  }

  return {
    listReports,
    getReport,
    addReport,
    updateReport,
    deleteReport
  };
}

module.exports = {
  createReportStore
};
