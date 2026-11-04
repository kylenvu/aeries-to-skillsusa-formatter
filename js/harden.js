// ============================================================
// Security hardening. Loaded before any other script.
// ============================================================
//
// Freezing Object.prototype blocks "prototype pollution": a specially
// crafted spreadsheet that tries to add properties to every object in
// the page while it is being read. With the prototype frozen, those
// writes are ignored. Nothing in this app needs to change it.
"use strict";
Object.freeze(Object.prototype);
