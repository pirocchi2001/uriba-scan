#!/usr/bin/env node
/**
 * 転送.xlsx を暗号化して master.enc を作る（公開してよいのは master.enc だけ）。
 *
 *   MASTER_PASSWORD='ログインパスワード' node tools/encrypt-master.js ~/Downloads/転送.xlsx
 *
 * パスワードを変える場合も、新しいパスワードでこれを実行して master.enc を差し替える。
 * KDF_SALT / KDF_ITERATIONS は app.js と同じ値にすること。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('../vendor/xlsx.core.min.js');
const parseMaster = require('../master-parse.js');

const KDF_SALT = 'uriba-scan-master-v1';
const KDF_ITERATIONS = 600000;

const src = process.argv[2];
const password = process.env.MASTER_PASSWORD;
if (!src || !password) {
  console.error("使い方: MASTER_PASSWORD='パスワード' node tools/encrypt-master.js <転送.xlsx>");
  process.exit(1);
}

const items = parseMaster(XLSX, fs.readFileSync(src));
const count = Object.keys(items).length;
if (count === 0) {
  console.error('JANコードが見つかりませんでした（D列=JAN, E列=品番名, F列=在売価）。');
  process.exit(1);
}

const itemsJson = JSON.stringify(items);
const payload = {
  version: crypto.createHash('sha256').update(itemsJson).digest('hex').slice(0, 16),
  file: path.basename(src),
  updatedAt: fs.statSync(src).mtimeMs,
  items,
};

const key = crypto.pbkdf2Sync(password, KDF_SALT, KDF_ITERATIONS, 32, 'sha256');
const iv = crypto.randomBytes(12);
const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
const enc = Buffer.concat([cipher.update(JSON.stringify(payload), 'utf8'), cipher.final(), cipher.getAuthTag()]);

const out = path.join(__dirname, '..', 'master.enc');
fs.writeFileSync(out, JSON.stringify({ v: 1, iv: iv.toString('base64'), data: enc.toString('base64') }));
console.log(`master.enc を作成しました（${count}件, 版 ${payload.version}）`);
