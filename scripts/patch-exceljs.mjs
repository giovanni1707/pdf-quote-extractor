import fs from 'fs';
import path from 'path';

function patchMinified(filePath) {
  if (!fs.existsSync(filePath)) return;
  let code = fs.readFileSync(filePath, 'utf8');
  let changed = false;

  if (code.includes('e.comments=t.comments[n.Target].comments')) {
    code = code.replace(
      /e\.comments=t\.comments\[n\.Target\]\.comments/g,
      'e.comments=(t.comments&&t.comments[n.Target]&&t.comments[n.Target].comments)?t.comments[n.Target].comments:[]'
    );
    changed = true;
  }

  if (code.includes('t.vmlDrawings[n.Target].comments')) {
    code = code.replace(
      /t\.vmlDrawings\[n\.Target\]\.comments/g,
      '(t.vmlDrawings&&t.vmlDrawings[n.Target]&&t.vmlDrawings[n.Target].comments?t.vmlDrawings[n.Target].comments:[])'
    );
    changed = true;
  }

  if (code.includes('e.note=Object.assign({},e.note,r[t])')) {
    code = code.replace(
      /e\.note=Object\.assign\(\{\},e\.note,r\[t\]\)/g,
      'e.note=Object.assign({},e.note,r&&r[t]?r[t]:{})'
    );
    changed = true;
  }

  // Also check formatted variants (spaces around operators)
  if (code.includes('(e.comments = t.comments[n.Target].comments)')) {
    code = code.replace(
      /\(e\.comments = t\.comments\[n\.Target\]\.comments\)/g,
      '(e.comments = (t.comments && t.comments[n.Target] && t.comments[n.Target].comments) ? t.comments[n.Target].comments : [])'
    );
    changed = true;
  }

  if (code.includes('const r = t.vmlDrawings[n.Target].comments;')) {
    code = code.replace(
      /const r = t\.vmlDrawings\[n\.Target\]\.comments;/g,
      'const r = (t.vmlDrawings && t.vmlDrawings[n.Target] && t.vmlDrawings[n.Target].comments) ? t.vmlDrawings[n.Target].comments : [];'
    );
    changed = true;
  }

  if (code.includes('e.note = Object.assign({}, e.note, r[t]);')) {
    code = code.replace(
      /e\.note = Object\.assign\(\{\}, e\.note, r\[t\]\);/g,
      'e.note = Object.assign({}, e.note, r && r[t] ? r[t] : {});'
    );
    changed = true;
  }

  if (changed) {
    fs.writeFileSync(filePath, code, 'utf8');
    console.log(`Successfully patched minified: ${filePath}`);
  }
}

function patchUnminified(filePath) {
  if (!fs.existsSync(filePath)) return;
  let code = fs.readFileSync(filePath, 'utf8');
  let changed = false;

  // Pattern: model.comments = options.comments[rel.Target].comments;
  if (/model\.comments\s*=\s*options\.comments\[rel\.Target\]\.comments;/.test(code)) {
    code = code.replace(
      /model\.comments\s*=\s*options\.comments\[rel\.Target\]\.comments;/g,
      'model.comments = (options.comments && options.comments[rel.Target] && options.comments[rel.Target].comments) ? options.comments[rel.Target].comments : [];'
    );
    changed = true;
  }

  // Pattern: const vmlComment = options.vmlDrawings[rel.Target].comments;
  if (/const\s+vmlComment\s*=\s*options\.vmlDrawings\[rel\.Target\]\.comments;/.test(code)) {
    code = code.replace(
      /const\s+vmlComment\s*=\s*options\.vmlDrawings\[rel\.Target\]\.comments;/g,
      'const vmlComment = (options.vmlDrawings && options.vmlDrawings[rel.Target] && options.vmlDrawings[rel.Target].comments) ? options.vmlDrawings[rel.Target].comments : [];'
    );
    changed = true;
  }

  // Pattern: comment.note = Object.assign({}, comment.note, vmlComment[index]);
  if (/comment\.note\s*=\s*Object\.assign\(\{\},\s*comment\.note,\s*vmlComment\[index\]\);/.test(code)) {
    code = code.replace(
      /comment\.note\s*=\s*Object\.assign\(\{\},\s*comment\.note,\s*vmlComment\[index\]\);/g,
      'comment.note = Object.assign({}, comment.note, (vmlComment && vmlComment[index]) ? vmlComment[index] : {});'
    );
    changed = true;
  }

  if (changed) {
    fs.writeFileSync(filePath, code, 'utf8');
    console.log(`Successfully patched unminified: ${filePath}`);
  }
}

// 1. Patch minified files
patchMinified(path.resolve('node_modules/exceljs/dist/exceljs.min.js'));
patchMinified(path.resolve('node_modules/.vite/deps/exceljs.js'));

// 2. Patch unminified files
patchUnminified(path.resolve('node_modules/exceljs/dist/exceljs.js'));
patchUnminified(path.resolve('node_modules/exceljs/dist/exceljs.bare.js'));
patchUnminified(path.resolve('node_modules/exceljs/lib/xlsx/xform/sheet/worksheet-xform.js'));
patchUnminified(path.resolve('node_modules/exceljs/dist/es5/xlsx/xform/sheet/worksheet-xform.js'));

console.log('ExcelJS patch execution complete.');
