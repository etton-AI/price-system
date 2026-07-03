import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Sealos/容器部署需要 standalone 模式
  output: "standalone",
  // xlsx 作为外部依赖，不被打包进 webpack bundle
  serverExternalPackages: ["xlsx"],
};

export default nextConfig;
