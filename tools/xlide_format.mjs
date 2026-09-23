// Format VBA sources with XLIDE's own formatter, for tools/xlide_differential.py.
//
//     node tools/xlide_format.mjs <bundle.mjs> <requests.json> <results.json>
//
// The bundle is XLIDE's src/analyzer/format/formatModule.ts built with the
// XLIDE checkout's esbuild (the Python side does that). Each request is
// {"id", "source"}; each result is {"id", "text"} or {"id", "refusal"}.
import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

const [bundlePath, requestsPath, resultsPath] = process.argv.slice(2);
const { formatVbaModule } = await import(pathToFileURL(bundlePath).href);
const requests = JSON.parse(readFileSync(requestsPath, 'utf8'));
const results = requests.map(({ id, source }) => {
  try {
    const result = formatVbaModule(source, { tabSize: 4, insertSpaces: true });
    return result.text === undefined ? { id, refusal: result.refusal ?? 'declined' } : { id, text: result.text };
  } catch (err) {
    return { id, error: String(err && err.stack ? err.stack : err) };
  }
});
writeFileSync(resultsPath, JSON.stringify(results));
