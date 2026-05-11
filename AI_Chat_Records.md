# AI 助手对话记录汇总

> **说明**：本项目研发过程中使用了 AI 助手进行辅助编程、Bug 排查及 Prompt 优化。

---

## 任务一：基于大模型的中文病例实体抽取

### 1. 基础代码框架构建
*   **对话链接**：https://claude.ai/share/add0c421-0108-4d8b-a62c-6bb3276a5cf3
*   **说明**：通过 AI 学习了如何使用 PyMuPDF 读取 PDF 内容，并初步构建了 Python 脚本的基础结构。

### 2. API 调用与 JSON 解析优化
*   **对话链接**：https://claude.ai/share/423b6313-749d-423c-9c49-ac396fa1e767
*   **说明**：向 AI 咨询了如何调用阿里云百炼的 OpenAI 兼容接口，并解决了模型返回结果中带有 Markdown 代码块导致 JSON 解析报错的问题。

### 3. 环境依赖整理与提交规范
*   **对话链接**：https://claude.ai/share/423b6313-749d-423c-9c49-ac396fa1e767
*   **说明**：请教 AI 如何编写 requirements.txt 以及如何按照指定的 GitHub 目录结构进行提交。

---

## 任务二：心衰患者生存预测与 LaTeX 小型英文论文撰写
### 1. 开发环境搭建与路径故障排查
*   **对话链接**：https://claude.ai/share/0a286ace-e693-4ec7-91b3-e4330e907a25
*   **说明**：在项目启动阶段，遇到了 Windows 路径转义导致的 SyntaxError 以及多版本 Python 导致的库缺失问题。通过向 AI 请教，学习了使用 r"" 原始字符串处理路径，并掌握了在 VS Code 中切换 Python 解释器（从 3.13 切换至已安装库的 3.7 环境）的方法。

### 2. 数据预处理与描述性统计可视化
*   **对话链接**：https://claude.ai/share/fdbed833-ec31-4ddd-aa4b-e768d2eda070
                 https://claude.ai/share/cb3bc581-98c8-4c63-8a37-0bef8a84d000
*   **说明**：针对原始数据集进行了缺失值检查及标准化处理。通过 AI 辅助编写了集成化的可视化脚本，生成了包含目标分布、异常值检测及特征分布的综合面板，并根据 AI 建议优化了绘图布局（Layout），解决了子图重叠和字体模糊的问题。

### 3. 因子检测与相关性热力图深度分析
*   **对话链接**：https://claude.ai/share/0144703e-1679-4226-a7f0-9aecdd5b2a03
                 https://claude.ai/share/fdbed833-ec31-4ddd-aa4b-e768d2eda070
*   **说明**：通过 AI 学习了如何解读相关性热力图，识别出 time、serum_creatinine 和 ejection_fraction 为核心影响因子。在 AI 的指导下，通过调整 annot_kws 和 rotation 参数，解决了热力图数值看不清及坐标轴标签遮挡的技术细节。

### 4. 特征重要性提取与医学逻辑验证
*   **对话链接**：https://claude.ai/share/50fab346-3c6f-4406-befb-a12b8e3d9297
                 https://claude.ai/share/c8605f4e-5542-4996-8bca-8a7a06f6ae9f
*   **说明**：向 AI 咨询了“相关性”与“重要性”的区别，解释了为何某些指标虽有统计关联但在预测模型中贡献有限。通过 AI 编写了基于随机森林（Random Forest）的特征重要性排序代码，为后续论文的“讨论”章节提供了医学逻辑支撑。

### 4. latex论文撰写与图表绘制
*   **对话链接**：https://claude.ai/share/c344b881-783c-428e-862d-1beba98ff4c9
*   https://claude.ai/share/78d730e7-d4c6-471a-ba23-04403866aee3
*  https://claude.ai/share/af57d68c-df89-4e2b-bb1e-4cc806d2756e
*   https://claude.ai/share/b7229da7-e20f-47fd-af12-e4b7ec698ba3

---

## 任务三：医疗大模型新手指南撰写
*   **对话链接**：https://claude.ai/share/c6c3b678-3c6a-41e7-a0ae-259d49d164d2
*   https://claude.ai/share/4f890268-19e9-4c22-a9d5-7da484e2b42b
