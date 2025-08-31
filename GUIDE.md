## Dify 插件开发用户指南

你好！看起来你已经创建了一个插件，现在让我们开始开发吧！

### 选择你要开发的插件类型

在开始之前，你需要了解一些关于插件类型的基础知识。插件支持在 Dify 中扩展以下功能：
- **工具（Tool）**：工具提供商，如 Google Search、Stable Diffusion 等，可用于执行特定任务。
- **模型（Model）**：模型提供商，如 OpenAI、Anthropic 等，你可以使用它们的模型来增强 AI 能力。
- **端点（Endpoint）**：类似于 Dify 中的服务 API 和 Kubernetes 中的 Ingress，你可以将 HTTP 服务扩展为端点，并使用自己的代码控制其逻辑。

基于你想要扩展的能力，我们将插件分为三种类型：**工具（Tool）**、**模型（Model）** 和 **扩展（Extension）**。

- **工具（Tool）**：这是一个工具提供商，但不仅限于工具，你也可以在其中实现端点。例如，如果你正在构建一个 Discord 机器人，需要同时具备"发送消息"和"接收消息"功能，那么**工具**和**端点**都是必需的。
- **模型（Model）**：仅仅是一个模型提供商，不允许扩展其他功能。
- **扩展（Extension）**：其他时候，你可能只需要一个简单的 HTTP 服务来扩展功能，**扩展**是你的正确选择。

我相信你在创建插件时已经选择了正确的类型，如果没有，你可以稍后通过修改 `manifest.yaml` 文件来更改它。

### 清单文件（Manifest）

现在你可以编辑 `manifest.yaml` 文件来描述你的插件，以下是其基本结构：

- version(version, 必需)：插件版本
- type(type, 必需)：插件类型，目前仅支持 `plugin`，未来支持 `bundle`
- author(string, 必需)：作者，这是市场中的组织名称，也应该等于仓库的所有者
- label(label, 必需)：多语言名称
- created_at(RFC3339, 必需)：创建时间，市场要求创建时间必须小于当前时间
- icon(asset, 必需)：图标路径
- resource (object)：要申请的资源
  - memory (int64)：最大内存使用量，主要与 SaaS 无服务器资源申请相关，单位字节
  - permission(object)：权限申请
    - tool(object)：反向调用工具权限
      - enabled (bool)
    - model(object)：反向调用模型权限
      - enabled(bool)
      - llm(bool)
      - text_embedding(bool)
      - rerank(bool)
      - tts(bool)
      - speech2text(bool)
      - moderation(bool)
    - node(object)：反向调用节点权限
      - enabled(bool) 
    - endpoint(object)：允许注册端点权限
      - enabled(bool)
    - app(object)：反向调用应用权限
      - enabled(bool)
    - storage(object)：申请持久存储权限
      - enabled(bool)
      - size(int64)：允许的最大持久内存，单位字节
- plugins(object, 必需)：插件扩展特定能力的 yaml 文件列表，插件包中的绝对路径，如果需要扩展模型，需要定义一个类似 openai.yaml 的文件，并在此处填写路径，路径上的文件必须存在，否则打包将失败。
  - 格式
    - tools(list[string]): 扩展的工具供应商，详细格式请参考 [工具指南](https://docs.dify.ai/plugins/schema-definition/tool)
    - models(list[string])：扩展的模型供应商，详细格式请参考 [模型指南](https://docs.dify.ai/plugins/schema-definition/model)
    - endpoints(list[string])：扩展的端点供应商，详细格式请参考 [端点指南](https://docs.dify.ai/plugins/schema-definition/endpoint)
  - 限制
    - 不允许同时扩展工具和模型
    - 不允许没有扩展
    - 不允许同时扩展模型和端点
    - 目前每种扩展类型最多只支持一个供应商
- meta(object)
  - version(version, 必需)：清单格式版本，初始版本 0.0.1
  - arch(list[string], 必需)：支持的架构，目前仅支持 amd64 arm64
  - runner(object, 必需)：运行时配置
    - language(string)：目前仅支持 python
    - version(string)：语言版本，目前仅支持 3.12
    - entrypoint(string)：程序入口，在 python 中应该是 main

### 安装依赖

- 首先，你需要一个 Python 3.11+ 环境，因为我们的 SDK 需要这个版本。
- 然后，安装依赖：
    ```bash
    pip install -r requirements.txt
    ```
- 如果你想添加更多依赖，可以将它们添加到 `requirements.txt` 文件中。一旦你在 `manifest.yaml` 文件中将运行器设置为 python，`requirements.txt` 将自动生成并用于打包和部署。

### 实现插件

现在你可以开始实现你的插件了，通过以下示例，你可以快速了解如何实现自己的插件：

- [OpenAI](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/openai)：模型提供商的最佳实践
- [Google Search](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/google)：工具提供商的简单示例
- [Neko](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/neko)：端点组的有趣示例

### 测试和调试插件

你可能已经注意到插件根目录中有一个 `.env.example` 文件，只需将其复制为 `.env` 并填入相应的值。如果你想在本地调试插件，需要设置一些环境变量。

- `INSTALL_METHOD`：将此设置为 `remote`，你的插件将通过网络连接到 Dify 实例。
- `REMOTE_INSTALL_URL`：来自你的 Dify 实例的调试主机和 plugin-daemon 服务端口的 URL，例如 `debug.dify.ai:5003`。可以使用 [Dify SaaS](https://debug.dify.ai) 或 [自托管 Dify 实例](https://docs.dify.ai/en/getting-started/install-self-hosted/readme)。
- `REMOTE_INSTALL_KEY`：你应该从使用的 Dify 实例获取调试密钥，在插件管理页面的右上角，你可以看到一个带有 `debug` 图标的按钮，点击它就能获得密钥。

运行以下命令启动你的插件：

```bash
python -m main
```

刷新你的 Dify 实例页面，现在你应该能在列表中看到你的插件，但它会被标记为 `debugging`，你可以正常使用它，但不建议用于生产环境。

### 发布和更新插件

为了简化插件更新工作流程，你可以配置 GitHub Actions，在创建发布时自动向 Dify 插件仓库创建 PR。

##### 前提条件

- 你的插件源代码仓库
- dify-plugins 仓库的分叉
- 在你的分叉中有正确的插件目录结构

#### 配置 GitHub Action

1. 创建一个对你的分叉仓库具有写权限的个人访问令牌
2. 在你的源仓库设置中将其添加为名为 `PLUGIN_ACTION` 的密钥
3. 在 `.github/workflows/plugin-publish.yml` 创建工作流文件

#### 使用方法

1. 更新你的代码和 `manifest.yaml` 中的版本
2. 在你的源仓库中创建发布
3. 该操作会自动打包你的插件并向你的分叉仓库创建 PR

#### 优势

- 消除手动打包和 PR 创建步骤
- 确保发布过程的一致性
- 在频繁更新期间节省时间

---

有关详细设置说明和示例配置，请访问：[GitHub Actions 工作流文档](https://docs.dify.ai/plugins/publish-plugins/plugin-auto-publish-pr)

### 打包插件

最后，只需运行以下命令打包你的插件：

```bash
dify-plugin plugin package ./ROOT_DIRECTORY_OF_YOUR_PLUGIN
```

你将得到一个 `plugin.difypkg` 文件，就这样，现在你可以将其提交到市场了，期待你的插件上架！


## 用户隐私政策

如果你想在市场上发布插件，请填写插件的隐私政策，详情请参考 [PRIVACY.md](PRIVACY.md)。