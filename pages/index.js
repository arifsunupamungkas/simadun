// pages/index.js
// Halaman utama — render full SPA dari public/index.html
// Di-serve sebagai static HTML via Next.js

import Head from 'next/head';

export default function Home() {
  return (
    <>
      <Head>
        <meta charSet="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Digitalisasi Kearsipan</title>
      </Head>
      {/* App di-mount via public/index.html yang di-redirect */}
    </>
  );
}

// Redirect ke static HTML
export async function getServerSideProps({ res }) {
  res.writeHead(302, { Location: '/app' });
  res.end();
  return { props: {} };
}
