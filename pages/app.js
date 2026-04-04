// pages/app.js
// Serve the full HTML SPA

import fs from 'fs';
import path from 'path';

export default function App() {
  return null; // rendered server-side
}

export async function getServerSideProps({ res }) {
  const filePath = path.join(process.cwd(), 'public', 'app', 'index.html');
  const html = fs.readFileSync(filePath, 'utf8');
  res.setHeader('Content-Type', 'text/html');
  res.write(html);
  res.end();
  return { props: {} };
}
