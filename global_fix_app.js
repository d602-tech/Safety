const fs = require('fs');
const path = require('path');

const filePath = 'd:/AI/GI01SafetyWalk/frontend/js/app.js';
let content = fs.readFileSync(filePath, 'utf8');

// Global replacement of corrupted template literal sequences
// 1. \${ -> ${
content = content.split('\\${').join('${');
// 2. \` -> `
content = content.split('\\`').join('`');

fs.writeFileSync(filePath, content, 'utf8');
console.log('Global cleanup of backslashes in app.js completed');
