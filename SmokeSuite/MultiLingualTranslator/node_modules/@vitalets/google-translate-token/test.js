import test from 'ava';
import puppeteer from 'puppeteer';
import {get as getToken} from './index';

test('check if what we generate equals to what translate.google.com generates', async t => {
    const browser = await puppeteer.launch();
    try {
        const token = await getToken('hello');
        const page = await browser.newPage();
        await page.goto('https://translate.google.com', {timeout: 10000, waitUntil: 'networkidle2'});
        const pRequest = page.waitForRequest(request => request.url().indexOf('tk=') > 0);
        await page.type('#source', 'hello');
        const request = await pRequest;
        const realToken = request.url().match(/tk=(\d+\.\d+)/i)[1];
        t.is(token.value, realToken);
    } catch (err) {
        t.fail(err);
    } finally {
        browser.close();
    }
});
