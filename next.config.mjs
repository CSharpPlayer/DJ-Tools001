/** @type {import('next').NextConfig} */
const basePath = process.env.NEXT_PUBLIC_BASE_PATH || "";

const nextConfig = {
  assetPrefix: basePath ? `${basePath}/` : undefined,
  basePath: basePath || undefined,
  output: "export",
  reactStrictMode: true
};

export default nextConfig;
