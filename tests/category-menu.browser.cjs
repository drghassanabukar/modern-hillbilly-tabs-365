const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const ts = require('typescript');
const { chromium } = require('playwright');

function readSdkPostbackGuard() {
  const file = require.resolve('@microsoft/sp-webpart-base/lib-commonjs/core/classicPages/ClassicPageUtils.js');
  const source = fs.readFileSync(file, 'utf8');
  const parsed = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  let implementation;
  function visit(node) {
    if (ts.isBinaryExpression(node) && node.left.getText(parsed).endsWith('.disableAutomaticPostbacks') && ts.isFunctionExpression(node.right)) {
      implementation = node.right.getText(parsed);
    }
    ts.forEachChild(node, visit);
  }
  visit(parsed);
  assert.ok(implementation, 'installed SDK postback guard must be available for the regression');
  return implementation;
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1000, height: 800 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const sourceRoot = path.join(__dirname, '../src/webparts/modernHillbillyTabs');
    const modules = {};
    for (const name of ['CategoryMenu', 'TabAppearance', 'CanvasTabs']) {
      modules[name] = ts.transpileModule(fs.readFileSync(path.join(sourceRoot, name + '.ts'), 'utf8'), {
        compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2018 }
      }).outputText;
    }
    const css = fs.readFileSync(path.join(sourceRoot, 'Tabs.css'), 'utf8');
    const sdkPostbackGuard = readSdkPostbackGuard();
    const scenarios = [
      { layout: 'menu', count: 24, guard: 'none', expectedMenu: true },
      { layout: 'auto', count: 24, guard: 'none', expectedMenu: true },
      { layout: 'auto', count: 8, guard: 'none', expectedMenu: false },
      { layout: 'tabs', count: 24, guard: 'none', expectedMenu: false },
      { layout: 'menu', count: 2, guard: 'sdk-capture', expectedMenu: true },
      { layout: 'auto', count: 9, guard: 'sdk-capture', expectedMenu: true },
      { layout: 'menu', count: 24, guard: 'sdk-capture', expectedMenu: true },
      { layout: 'auto', count: 24, guard: 'sdk-capture', expectedMenu: true },
      { layout: 'menu', count: 24, guard: 'editor-bubble', expectedMenu: true },
      { layout: 'auto', count: 24, guard: 'editor-bubble', expectedMenu: true }
    ];
    for (const scenario of scenarios) {
      await page.setContent(`<style>body{margin:24px;font:16px Segoe UI,sans-serif}section{max-width:760px}${css}</style><form id="page-form"><section id="canvas" data-automation-id="CanvasSection"><div data-automation-id="CanvasControl"><div data-sp-feature-instance-id="menu-host"><div id="menu-host"></div></div></div>${Array.from({ length: scenario.count }, (_, i) => `<div data-automation-id="CanvasControl"><div data-sp-feature-instance-id="part-${i}">Content ${i + 1}</div></div>`).join('')}</section></form>`);
      await page.evaluate(({ modules, scenario, sdkPostbackGuard }) => {
        const cache = {};
        function load(name) {
          name = name.replace('./', '');
          if (cache[name]) return cache[name].exports;
          const module = { exports: {} };
          cache[name] = module;
          new Function('require', 'exports', 'module', modules[name])(load, module.exports, module);
          return module.exports;
        }
        window.canvasClicks = 0;
        window.formSubmits = 0;
        document.getElementById('page-form').addEventListener('submit', event => { window.formSubmits++; event.preventDefault(); });
        const host = document.getElementById('menu-host');
        window.controller = new (load('CanvasTabs').CanvasTabs)(host,
          Array.from({ length: scenario.count }, (_, i) => ({ WebPartID: 'part-' + i, TabLabel: 'Category ' + String(i + 1).padStart(2, '0') })),
          { tabLayout: scenario.layout, tabDesign: 'outlined' }, 'browser-regression');
        if (scenario.guard === 'sdk-capture') {
          const sp_core_library_1 = { Validate: { isNotNullOrUndefined() {} }, EnvironmentType: { ClassicSharePoint: 1 } };
          const guard = new Function('sp_core_library_1', 'return (' + sdkPostbackGuard + ')')(sp_core_library_1);
          // Replay the real SDK's default-prevention pattern, without claiming all modern canvases use this classic-page hook.
          guard(host, 1);
        } else if (scenario.guard === 'editor-bubble') {
          document.getElementById('canvas').addEventListener('click', event => { window.canvasClicks++; event.preventDefault(); });
        }
      }, { modules, scenario, sdkPostbackGuard });
      const name = `${scenario.layout}/${scenario.count}/${scenario.guard}`;
      const details = page.locator('.hillbilly-category-details');
      assert.equal(await details.count(), scenario.expectedMenu ? 1 : 0, name + ': layout selection');
      if (scenario.expectedMenu) {
        const summary = page.locator('.hillbilly-category-summary');
        const search = page.getByRole('searchbox', { name: 'Find a category' });
        await summary.click();
        assert.equal(await details.evaluate(element => element.open), true, name + ': click opens the category chooser');
        await summary.click();
        assert.equal(await details.evaluate(element => element.open), false, name + ': click closes once without double toggling');
        await summary.press('Enter');
        await search.waitFor({ state: 'visible' });
        assert.equal(await details.evaluate(element => element.open), true, name + ': Enter opens');
        await search.fill('Category 02');
        assert.equal(await page.getByRole('tab').count(), 1, name + ': filtering shows one match');
        assert.equal(await page.locator('[data-sp-feature-instance-id="part-0"]').isVisible(), true, name + ': search preserves content');
        await search.press('ArrowDown');
        await page.getByRole('tab', { name: 'Category 02', exact: true }).press('Enter');
        assert.equal(await details.evaluate(element => element.open), false, name + ': selection closes');
        assert.match(await summary.textContent(), /Category 02/, name + ': summary follows selection');
        assert.equal(await summary.evaluate(element => element === document.activeElement), true, name + ': selection restores focus');
        assert.equal(await page.locator('[data-sp-feature-instance-id="part-1"]').isVisible(), true, name + ': selected content is visible');
        await summary.press('Space');
        assert.equal(await details.evaluate(element => element.open), true, name + ': Space opens once');
        await search.fill('No matching category');
        await search.press('Enter');
        assert.equal(await details.evaluate(element => element.open), true, name + ': empty Enter keeps chooser open');
        await search.press('Escape');
        assert.equal(await details.evaluate(element => element.open), false, name + ': Escape closes');
        assert.equal(await summary.evaluate(element => element === document.activeElement), true, name + ': Escape restores focus');
        assert.equal(await page.evaluate(() => window.formSubmits), 0, name + ': navigation never submits the page form');
        if (scenario.guard === 'editor-bubble') assert.equal(await page.evaluate(() => window.canvasClicks), 0, name + ': category interactions stay inside the web part');
      }
      await page.evaluate(() => window.controller.dispose());
      assert.equal(await page.locator('[data-sp-feature-instance-id^="part-"]:visible').count(), scenario.count, name + ': disposal restores content');
      console.log('Passed ' + name);
    }
    assert.deepEqual(errors, [], 'browser must not report runtime errors');
    console.log('Category menu browser regressions passed: native and guarded clicks, Enter/Space, search, focus, selection, thresholds, and cleanup.');
  } finally { await browser.close(); }
}

run().catch(error => { console.error(error); process.exitCode = 1; });
