const os = require('os');
const crypto = require('crypto');

// Generate a word code similar to how railway CLI does it
// The d parameter is base64 encoded: wordCode=XXX&hostname=YYY
const hostname = os.hostname();
// Generate random word code
const words = ['crimson', 'azure', 'golden', 'silver', 'verdant', 'cobalt', 'amber', 'violet', 'scarlet', 'teal'];
const adj = ['swift', 'bright', 'quiet', 'brave', 'bold', 'calm', 'wise', 'noble', 'keen', 'fair'];
const noun = ['river', 'mountain', 'forest', 'ocean', 'desert', 'valley', 'island', 'canyon', 'glacier', 'meadow'];

const wordCode = `${words[Math.floor(Math.random() * words.length)]}-${adj[Math.floor(Math.random() * adj.length)]}-${noun[Math.floor(Math.random() * noun.length)]}`;
const data = `wordCode=${wordCode}&hostname=${hostname}`;
const encoded = Buffer.from(data).toString('base64');
const url = `https://railway.com/cli-login?d=${encoded}`;

console.log('=== RAILWAY CLI LOGIN ===');
console.log('URL:', url);
console.log('Pairing code:', wordCode);
console.log('Encoded data:', data);
