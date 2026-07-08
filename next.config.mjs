/** @type {import('next').NextConfig} */
const isGitHubPages = process.env.GITHUB_PAGES === "true";
const githubPagesBasePath = "/DJ-Tools001";

const nextConfig = {
  assetPrefix: isGitHubPages ? `${githubPagesBasePath}/` : undefined,
  basePath: isGitHubPages ? githubPagesBasePath : undefined,
  output: "export",
  reactStrictMode: true
};

export default nextConfig;
