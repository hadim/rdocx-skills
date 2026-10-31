// Make a generator's output reproducible byte for byte: freeze the clock (core.xml dates and
// the zip entries' times; run node with TZ=UTC) and seed Math.random and crypto.getRandomValues
// (random ids). Require this first, before docx or pptxgenjs. Nothing else is changed.
"use strict";

const FIXED_MS = Date.UTC(2026, 0, 15, 9, 0, 0);
const RealDate = Date;
class FixedDate extends RealDate {
	constructor(...args) {
		if (args.length === 0) super(FIXED_MS);
		else super(...args);
	}
	static now() {
		return FIXED_MS;
	}
}
globalThis.Date = FixedDate;

// mulberry32
let state = 0x5eed1234 >>> 0;
function rand() {
	state = (state + 0x6d2b79f5) >>> 0;
	let t = state;
	t = Math.imul(t ^ (t >>> 15), t | 1);
	t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
	return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
}
Math.random = rand;
if (globalThis.crypto) {
	globalThis.crypto.getRandomValues = (arr) => {
		for (let i = 0; i < arr.length; i++) arr[i] = Math.floor(rand() * 256);
		return arr;
	};
}

module.exports = { FIXED_MS };
