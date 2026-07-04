# ---- Build Stage ----
FROM node:22-alpine AS builder
WORKDIR /app

# 安装依赖
COPY package.json package-lock.json* ./
RUN npm ci

# 构建
COPY . .
RUN mkdir -p /app/public/data /app/data

# 生成价格数据 (解析 Excel → JSON)
RUN node parsers/build_db.js

RUN npm run build

# ---- Production Stage ----
FROM node:22-alpine AS runner
WORKDIR /app

ENV NODE_ENV=production
ENV NEXT_TELEMETRY_DISABLED=1
# 限制 Node.js 最大老生代堆内存为 768MB（容器 1.5Gi 限制内安全值）
ENV NODE_OPTIONS="--max-old-space-size=768"

# 只复制 standalone 产出
COPY --from=builder /app/.next/standalone ./
COPY --from=builder /app/.next/static ./.next/static
COPY --from=builder /app/public ./public
COPY --from=builder /app/parsers ./parsers
COPY --from=builder /app/node_modules/xlsx ./node_modules/xlsx

# 保留一份构建时生成的数据作为初始种子（不会被 PVC 覆盖）
RUN mkdir -p /app/public/data-init /app/data-init
RUN cp /app/public/data/prices.json /app/public/data-init/prices.json 2>/dev/null || true
RUN cp /app/data/prices.json /app/data-init/prices.json 2>/dev/null || true

# 启动脚本：首次部署时，将构建数据复制到 PVC（如果 PVC 为空）
RUN printf '#!/bin/sh\n\
set -e\n\
if [ ! -f /app/public/data/prices.json ]; then\n\
  echo "[startup] PVC is empty, seeding from build cache..."\n\
  mkdir -p /app/public/data /app/data\n\
  cp /app/public/data-init/prices.json /app/public/data/prices.json\n\
  cp /app/data-init/prices.json /app/data/prices.json\n\
  echo "[startup] Seed done"\n\
else\n\
  echo "[startup] PVC has data, syncing to backup path..."\n\
  mkdir -p /app/data\n\
  cp /app/public/data/prices.json /app/data/prices.json 2>/dev/null || true\n\
fi\n\
echo "[startup] Starting Next.js..."\n\
exec node server.js\n\
' > /app/start.sh && chmod +x /app/start.sh

EXPOSE 3000
CMD ["/app/start.sh"]
