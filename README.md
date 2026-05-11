# 医疗大模型研究组 

> **仓库作者：** zhangmuhan
> **最后更新：** 2026-05

---

## 项目概述

本项目为医疗大模型研究组新人中期考核的成果提交仓库。项目涵盖了从 LLM API 调用、临床数据挖掘与机器学习建模、LaTeX 学术论文撰写到知识沉淀总结的全流程科研任务。本项目严格按照任务书规范进行目录划分，包含三个任务的完整代码、数据和文档，分别放在 TASK1/、TASK2/、TASK3/ 三个目录下。


## 项目概述仓库结构与提交物清单
zhangmuhan/
│
├── TASK1/                          # 基于大模型的中文病例实体抽取
│   ├── entity_extraction.py        # 主脚本：API 调用与实体抽取逻辑
│   ├── requirements.txt            # Python 依赖列表
│   ├── cases/                      # 原始 PDF 病例文件
│   │   └── *.pdf
│   └── results/                    # 抽取结果（JSON 格式）
│       └── *.json
│
├── TASK2/                          # 心衰患者生存预测与学术论文
│   ├── analysis.py      # 主脚本：数据处理与模型训练
│   ├── requirements.txt            # Python 依赖列表
│   ├── dataset/                    # 心衰数据集
│   │   └── heart_failure.csv
│   ├── paper/                      # LaTeX 论文源文件
│   │   ├── main.tex
│   │   └── *.bib
│   └── paper_output.pdf            # 最终生成的 PDF 论文
│
├── TASK3/                          # 医疗大模型新人快速上手指南
│   ├── 医疗大模型新人快速上手指南.md         # Markdown 版指南
│   └── 医疗大模型新人快速上手指南.pdf        # PDF 版指南
│
├── AI_Chat_Records.md              # AI 对话链接汇总
└── README.md                       # 本文件


---

## 任务要点：
## TASK 1：基于大模型的中文病例实体抽取
实现功能： 使用 Python 调用大模型 API（Qwen），通过设计结构化 Prompt 从 PDF 病例中提取患者基本信息、症状、诊断及治疗方案。
核心产出： 实现了医疗实体到标准化 JSON 格式的自动化转化。
核心抽取实体类型包括患者基本信息、主要症状、既往史、诊断结果、治疗方案等

---

## TASK 2：心衰患者生存预测与学术论文数据分析
对心衰临床数据集进行了描述性统计、相关性热力图分析及特征重要性筛选（Top 3 风险因子）。模型构建时对比了多种机器学习模型（如 Logistic Regression, Random Forest），并进行了四个维度的模型性能评估。采用 LaTeX 编写了完整学术论文，包含公式推导、基线特征表及 AI 生成的科技感流程图。
其中机器学习全流程包含：

#### 1 数据清洗与预处理
- 缺失值检测
- 异常值识别（IQR 方法、3σ 原则）
- 数据集划分（训练集 / 验证集 / 测试集，比例 8：2）

#### 2 特征工程与重要性分析
- 相关性热力图分析（Pearson ）
- 特征重要性排序（基于 Random Forest）

#### 3 模型训练与对比
- 基线模型：Logistic Regression
- 集成模型：Random Forest、XGBoost

#### 4 模型评估
- 评估指标：**AUC-ROC、F1-Score、Accuracy、Sensitivity、Specificity**
- 可视化：ROC 曲线、混淆矩阵


## TASK 3：医疗大模型新人快速上手指南
本指南是基于自身伸展经验撰写的一份避坑手册，记录了作者在加入医疗大模型研究组初期遇到的真实问题与解决方案，力求让后来的同学少走弯路。文档总结了 Python 环境管理（及大模型 Debug 技巧，旨在为医学背景零基础学弟学妹提供“无痛”上手方案。

Note: 本项目所有内容均通过 AI 协作生成并经过人工校验。详细的对话 Prompt 逻辑请参考 AI_Chat_Records.md。
