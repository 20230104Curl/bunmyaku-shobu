import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const read = (file) => fs.readFileSync(path.join(root, file), 'utf8');

const index = read('index.html');
const app = read('assets/app.js');
const config = read('assets/config.js');
const code = read('apps-script/Code.gs');

assert.match(index, /<title>文脈勝負<\/title>/);
assert.match(index, /assets\/config\.js/);
assert.match(index, /assets\/app\.js/);
assert.match(config, /__APPS_SCRIPT_WEB_APP_URL__/);

new Function(app);
new Function(code);

assert.match(app, /readingSeconds: 120/);
assert.match(app, /totalSeconds: 180/);
assert.match(app, /durationLabel/);
assert.doesNotMatch(app, /3分間で|180秒後|4分間/);

for (const functionName of ['doGet', 'doPost', 'setupWorkbook', 'refreshDashboards', 'refreshIndividualSearch']) {
  assert.match(code, new RegExp('function\\s+' + functionName + '\\s*\\('));
}

const publicFiles = [
  'index.html', 'assets/app.js', 'assets/config.js', 'assets/styles.css',
  'apps-script/Code.gs', 'README.md', 'SECURITY.md'
].map(read).join('\n');

assert.equal(publicFiles.includes('調べることと「寄り道」'), false, '問題本文・タイトルを公開リポジトリへ置かない');
assert.equal(publicFiles.includes('A → D → B → E → F → C'), false, '正解順を公開リポジトリへ置かない');
assert.equal(publicFiles.includes('sanaru0104@gmail.com'), false, '管理者メールを公開リポジトリへ置かない');

console.log('文脈勝負 GitHub版の構成検査に合格しました。');
