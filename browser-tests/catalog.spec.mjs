import { test, expect } from '@playwright/test';

test('gerçek koleksiyon yüklenir, varyantla bulunan kişi kaynaklı gösterilir', async ({ page }) => {
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto('/');
  await expect(page.locator('#people-count')).toHaveText('7.934');
  await expect(page.locator('.person-card')).toHaveCount(20);
  await page.locator('#search-input').fill('İbni Sina');
  await page.locator('#search-form').evaluate(form => form.requestSubmit());
  await expect(page.locator('.person-card').first()).toContainText('İbn Sînâ');
  await page.locator('.person-card').first().click();
  await expect(page.locator('#detail-name')).toHaveText('İbn Sînâ');
  await expect(page.locator('#detail-date')).toHaveText('(ö. 428/1037)');
  await expect(page.locator('#detail-source')).toHaveAttribute('href', 'https://islamansiklopedisi.org.tr/ibn-sina');
  await expect(page.locator('#citation-text')).toHaveValue('İbn Sînâ (ö. 428/1037)');
  await page.screenshot({ path: '.local/vefeyat-desktop.png' });
  expect(errors).toEqual([]);
});

test('aynı adlı iki şahıs ayrılır ve düzeltme formu doğru kaydı gösterir', async ({ page }) => {
  await page.goto('/');
  await page.locator('[data-search="İbn Rüşd"]').click();
  const matches = page.locator('.person-card').filter({ has: page.locator('.person-name', { hasText: /^İbn Rüşd$/ }) });
  await expect(matches).toHaveCount(2);
  await expect(matches.nth(0)).toContainText(/520\/1126|595\/1198/);
  await expect(matches.nth(1)).toContainText(/520\/1126|595\/1198/);
  await matches.nth(1).click();
  await expect(page.locator('#detail-name')).toHaveText('İbn Rüşd');
  await page.locator('#correct-person').click();
  await expect(page.locator('#contribution-dialog')).toBeVisible();
  await expect(page.locator('#contribution-heading')).toHaveText('Düzeltme önerin');
  await expect(page.locator('#contribution-name')).toHaveValue('İbn Rüşd');
  await expect(page.locator('#contribution-source')).toHaveValue(/islamansiklopedisi\.org\.tr\/ibn-rusd--/);
  await page.keyboard.press('Escape');
  await expect(page.locator('#contribution-dialog')).not.toBeVisible();
});

test('yakın eşleşme ve sonuçsuz arama anlaşılır biçimde gösterilir', async ({ page }) => {
  await page.goto('/');
  await page.locator('#search-input').fill('gazzli');
  await page.locator('#search-form').evaluate(form => form.requestSubmit());
  await expect(page.locator('#results-heading')).toHaveText('Yakın isim eşleşmeleri');
  await expect(page.locator('#result-count')).toContainText('Bunu mu demek istediniz?');
  await page.locator('#search-input').fill('zzzzzzzzzzzzzzzzzzzzz');
  await page.locator('#search-form').evaluate(form => form.requestSubmit());
  await expect(page.locator('#empty-state')).toBeVisible();
  await expect(page.locator('.person-card')).toHaveCount(0);
});

test('mobil görünüm yatay taşmaz ve kayıt ayrıntısı açılır', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page.locator('[data-search="İbn Sînâ"]').click();
  await expect(page.locator('.person-card').first()).toContainText('İbn Sînâ');
  await page.locator('.person-card').first().click();
  await expect(page.locator('#detail-content')).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBeTruthy();
  await page.screenshot({ path: '.local/vefeyat-mobile.png' });
});

test('Word dışında manuel arama çalışır, belgeye ekleme kapalı kalır', async ({ page }) => {
  // Test the intentional browser fallback without loading the Office host bridge.
  await page.route('https://appsforoffice.microsoft.com/**', route => route.fulfill({ contentType: 'text/javascript', body: '' }));
  await page.setViewportSize({ width: 380, height: 800 });
  await page.goto('/word/');
  await expect(page.locator('#get-selection')).toBeDisabled();
  await page.locator('#query').fill('İbni Sina');
  await page.locator('#search-form').evaluate(form => form.requestSubmit());
  await expect(page.locator('#results .person').first()).toContainText('İbn Sînâ');
  await page.locator('#results input[type=radio]').first().check();
  await expect(page.locator('#chosen-date')).toHaveText('(ö. 428/1037)');
  await expect(page.locator('#insert-date')).toBeDisabled();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBeTruthy();
});
