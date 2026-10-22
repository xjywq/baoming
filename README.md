# 报名表导出

把报名小程序导出的报名表（.xls）整理成按时段排序的预约名单（.xlsx）。

- 线上地址：https://baoming.djy.workers.dev
- `public/`：静态页面，表格转换完全在浏览器里完成
- `src/worker.js`：`/api/fetch` 代下载小程序导出链接（只放行 `https://api-xcx.qunsou.co/mini/enroll/v2/export?id=...`）

## 开发

```bash
npx wrangler dev --ip 0.0.0.0   # 本地预览，手机同网段可访问
node --test                     # 运行测试
```

## 部署

Cloudflare Workers Builds 关联本仓库，push 到 `main` 自动部署：

- 构建命令：`node --test`
- 部署命令：`npx wrangler deploy`
