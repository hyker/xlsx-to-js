// Resource-amplification regressions run in a killable, memory-limited process.
import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
import { XlsxParser } from '../dist/index.js';
import { fixture, worksheetXml } from './fixtures.mjs';

globalThis.DOMParser = new JSDOM('').window.DOMParser;
const bytes = await fixture({sheet:worksheetXml('',process.argv[2])});
await assert.rejects(new XlsxParser().readFile(bytes,{dense:process.argv[3]==='true'}),/reference|budget|limits/i);
