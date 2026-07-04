import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: 'export',
  basePath: '/Audio-to-text',
  allowedDevOrigins: ['192.168.1.155', 'localhost'],
};

export default nextConfig;
