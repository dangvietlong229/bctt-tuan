'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

/**
 * Replace a file only after the complete new contents have been written in the
 * same directory. Keeping the temporary file beside the destination makes the
 * final rename atomic on the filesystems used by the report workflow.
 */
function writeTextAtomic(filePath, text, validate) {
  if (typeof text !== 'string' || text.length === 0) {
    throw new Error(`Refusing to publish an empty file: ${filePath}`);
  }
  if (validate) validate(text);

  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true });
  const suffix = crypto.randomBytes(6).toString('hex');
  const tmpPath = path.join(dir, `.${path.basename(filePath)}.${process.pid}.${suffix}.tmp`);

  try {
    const fd = fs.openSync(tmpPath, 'wx');
    try {
      fs.writeFileSync(fd, text, 'utf8');
      fs.fsyncSync(fd);
    } finally {
      fs.closeSync(fd);
    }
    fs.renameSync(tmpPath, filePath);
  } catch (error) {
    try {
      fs.unlinkSync(tmpPath);
    } catch {
      // The temporary file may not have been created, or the rename succeeded.
    }
    throw error;
  }
}

function writeJsonAtomic(filePath, value, validate) {
  if (validate) validate(value);
  const text = `${JSON.stringify(value, null, 2)}\n`;
  writeTextAtomic(filePath, text, (candidate) => JSON.parse(candidate));
}

module.exports = { writeTextAtomic, writeJsonAtomic };
