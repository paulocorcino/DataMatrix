// JavaScript Document


var now = new Date();
// fix the bug in Navigator 2.0, Macintosh
fixDate(now);
now.setTime(now.getTime() + 1000 * 24 * 60 * 60 * 1000);
