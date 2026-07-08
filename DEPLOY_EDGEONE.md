# EdgeOne Makers / Pages 上线说明

## 推荐方式：直接上传静态产物

这个项目已配置为 Next.js 静态导出，适合使用 EdgeOne Makers / EdgeOne Pages 的免费方案部署。

1. 构建项目：

```powershell
pnpm.cmd build
```

2. 构建完成后，静态网站文件会生成在：

```text
out
```

3. 登录 EdgeOne Makers 控制台，创建 Pages 项目。

4. 选择直接上传，把 `out` 文件夹内的所有文件上传，或上传本项目根目录下生成的：

```text
edgeone-pages-static.zip
```

5. 发布完成后，EdgeOne 会分配一个可访问域名；如需绑定自己的域名，在项目设置里添加自定义域名并按提示配置 DNS。

## 如果使用 Git 导入

项目类型选择 Next.js 或静态站点均可，推荐配置：

```text
Install command: pnpm install --frozen-lockfile
Build command: pnpm build
Output directory: out
```

## 注意

- 当前工具的 PDF、Word 解析和 Word 导出都在浏览器端完成，不需要后端服务。
- `.doc` 旧格式仍需先另存为 `.docx` 后上传。
- 生成的会议记录不会上传到服务器，用户文件只在访问者自己的浏览器里处理。
