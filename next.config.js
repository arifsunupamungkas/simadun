/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  // Allow large payloads for file uploads
  experimental: {
    serverComponentsExternalPackages: []
  }
};

module.exports = nextConfig;
