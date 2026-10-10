# docx-jsx Agent 工作约定

本仓库的目标是用 Rust CLI 将 JSX/TSX 描述编译为 DOCX，并将外部 DOCX
逆向为可重新编译的 JSX。编译后端必须使用 `docx-rs`；不要引入 Node.js
作为运行时或编译后端。

## 核心不变量

- 对 `docx-rs` 公开且能生成合法 OOXML 的 writer API，持续提供 JSX/IR 封装。
- 正向与逆向功能必须 1:1：新增可写属性或组件时，同时实现外部 DOCX 逆向；
  不能仅依赖嵌入的 IR manifest 证明逆向可用。
- 逆向不得无提示地合并或丢弃结构、样式或数值。仅当语义完全相同时，才能
  折叠为简写；无法无损表示时返回包含原因和修复建议的错误。
- CLI 面向 Agent：命令和输出应确定、可脚本化；`spec` 输出机器/Agent 可读规范，
  `validate` 只校验 DSL，不写 DOCX。
- 保持增量构建和增量缓存；不得为了临时排错删除全局 Cargo 缓存或整个 `target`。

## 每个功能批次的流水线

1. 覆盖审计
   - 以当前锁定版本 `docx-rs` 的公开 writer API 和实际序列化结果为准。
   - 更新 `docs/docx-rs-coverage.md`，区分：已覆盖、上游 writer 缺失、上游输出
     不合法、reader-only。
   - 不把 reader-only API 或会生成 schema-invalid XML 的 setter 宣称为已支持。

2. Spec first
   - 先在 `docs/spec.md` 定义组件、属性、单位、默认值、互斥关系、合法嵌套和
     正反向行为。
   - 必要时同步 `spec/ir-v1.schema.json`；`docx-jsx spec` 的输出必须随规范更新。
   - API 名称应稳定、可组合，并尽量直接对应 OOXML/`docx-rs` 概念。

3. TDD 红灯
   - 先添加 IR 校验测试、编译 XML 测试和外部 `docx-rs` DOCX 逆向测试。
   - 运行定向测试，确认它因预期缺口失败，而不是测试本身错误。
   - 对无效输入测试错误路径、原因和可执行的修复建议。

4. 最小完整实现
   - 实现属性白名单和语义校验。
   - 实现 JSX/IR 到 `docx-rs` 的正向编译。
   - 实现不依赖 IR manifest 的外部 DOCX 到 JSX 逆向。
   - 对范围结构（书签、批注、修订等）按匹配 ID 重建嵌套，不得退化成相邻
     空标记。
   - 正反向共用单位转换、枚举和结构 helper，避免两套语义漂移。

5. 文档和集成夹具
   - 更新 README 功能矩阵、`docs/spec.md` 和 `docs/docx-rs-coverage.md`。
   - 将新能力加入 `tests/fixtures/openxml-valid.tsx` 或对应 CLI 夹具。
   - 示例不得引用开发机私有路径或不可发布的素材。

6. 本地质量门
   - `cargo fmt -- --check`
   - `cargo test --all-targets --all-features --locked`
   - `cargo clippy --all-targets --all-features --locked -- -D warnings`
   - `./scripts/openxml-test.sh`
   - `git diff --check`
   - Open XML 集成测试使用 .NET `DocumentFormat.OpenXml` 验证 Microsoft 365
     文件格式；不能只以 ZIP 可打开或 `docx-rs` 能回读作为合法性证据。

7. 远端验证
   - 本地验证通过后运行：
     `./scripts/gpu-test.sh --all-targets --all-features --locked`。
   - 脚本只同步源码，排除 `.git`、`target` 和私有素材，并在 `gpu-server`
     使用增量构建。
   - 远端暂时不可连接时，记录准确错误并继续本地功能开发，不等待 CI 或远端
     恢复；不得把未执行的远端测试报告为通过。

8. 提交与发布
   - 每个已经完成正向、逆向、规范和验证的功能批次形成独立提交。
   - 发布前更新 `Cargo.toml`、`Cargo.lock`、CHANGELOG 和必要的 README 版本信息。
   - 用 `dist plan --tag=vX.Y.Z` 检查 cargo-dist 清单。
   - 推送 main 后创建并推送带注释的版本标签；使用 `gh` 检查 Actions 和 GitHub
     Release 是否真实创建。标签存在不等于 Release 发布成功。
   - 不改写已经推送的版本标签；发布失败应修复原因并发布新的补丁版本。
   - 不阻塞等待 CI：流水线运行期间继续开发下一个组件，但必须准确记录发布
     状态，最终确认 Release 和跨平台资产存在。

## 镜像与网络

- Cargo/crates.io 优先使用仓库 `.cargo/config.toml` 中的 rsproxy 国内镜像。
- rustup 使用 `https://rsproxy.cn`，Rusty V8 使用仓库配置的国内代理。
- 对镜像瞬时抖动使用有限重试、较长超时和低速容忍；不要因一次超时立即切换
  回海外源，也不要绕过 `--locked` 发布。
- cargo-dist 生成的 workflow 不应直接手工修改为与生成清单不一致；通用网络设置
  放在 `.cargo/config.toml` 或 cargo-dist 支持的配置入口。

## 私有文件与工作树安全

- 用户提供的原始 DOCX 模板、逆向产物、视觉对比图片和临时解包目录不得提交。
- 特别禁止提交位于 Windows Downloads 的经济运行分析周报模板或其复制品。
- 真实模板实验使用 `/tmp` 目录，并在提交前通过 `git status` 和 `git ls-files`
  确认仓库中没有 `.docx` 私有素材。
- 工作树可能包含用户或上一批尚未完成的修改。保留无关改动，不使用
  `git reset --hard`、`git checkout --` 或其他破坏性清理命令。

## 完成判定

单个组件只有在规范、有效/无效 TDD、正向 XML、外部 DOCX 逆向、CLI/Agent
接口、README 矩阵和 Open XML SDK 校验全部完成后，才算该批次完成。整个目标
“封装 docx-rs 支持的全部功能”只有在覆盖审计逐项提供证据、无剩余公开 writer
API 缺口时才能宣告完成；不能用当前测试全绿替代完整覆盖证明。
