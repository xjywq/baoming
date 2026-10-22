# 报名表导出

把报名小程序导出的报名表（.xls）整理成按时段排序的预约名单（.xlsx）。

- 线上地址：https://xjywq.github.io/baoming/
- 纯静态网页，没有第三方库：`public/xls.js` 读 .xls（BIFF8），`public/xlsx.js` 写 .xlsx，转换全在浏览器里完成，名单不经过任何服务器
- 支持粘贴小程序导出链接（只放行 `https://api-xcx.qunsou.co/mini/enroll/v2/export?id=...`），该接口允许跨域，浏览器直接下载
- 分享按钮只在 iOS / Safari 显示：Chromium 内核不允许分享 .xlsx

## 开发

```bash
npm ci             # 安装测试依赖（SheetJS，只在测试里做对照，不进网站）
npm test           # 运行测试
npx serve public   # 本地预览
```

## 部署

push 到 `main` 后由 GitHub Actions（`.github/workflows/pages.yml`）跑测试并发布 `public/` 到 GitHub Pages。
