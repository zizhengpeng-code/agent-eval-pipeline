# 月末费用表复核

这是公开的合成教学题，当前状态为**待专家复核**。训练或比较已接触此仓库的模型时，应另用私有题目。

1. 阅读[考生题面](pack/task/task.md)，打开[费用表](pack/reference_files/benefits.xlsx)。三笔费用分别为 100、300 和 200 元。合计格的缓存显示 600，但求和区域包含合计格自身。
2. 查看[审稿记录](pack/review.csv)。第一版只问合计金额，照抄缓存即可答对。模拟审稿提出修改，第二版要求同时定位异常单元格。真人专家可以在当前版本填写结论、原因及修改要求。
3. 对照[参考答案](pack/ground_truth/answer.json)和[评分权重](pack/rubric/criteria.json)。地址集合完全匹配得 60 分，金额正确得 40 分。分别给分可以看出答题者漏掉了哪一步。

在仓库根目录运行：

```bash
python3 -m pip install -r requirements.txt
python3 examples/circular_audit/run.py
python3 examples/circular_audit/run.py --submission examples/circular_audit/submissions/cache_only.json
```

| 提交 | 定位自引用 | 重算金额 | 合计 |
| --- | ---: | ---: | ---: |
| 参考答案 | 60 | 40 | 100 |
| 仅抄缓存 | 0 | 40 | 40 |

每次运行都会核对[版本指纹](freeze.json)，调用盲测运行器的投放函数，检查考生目录只收到题面和工作簿，再用 Excel 重算引擎输出异常公式。评分明细列出每项的提交值、参考值和得分。评分脚本可直接读取真实运行留下的 workspace/output/answer.json。

示例指纹锁定了当前教学版本，含审稿记录、答案和评分权重。生产用多材料任务包按主仓库的 validate-task 和 freeze-task 流程验收。这道单工作簿练习只演示投放、指纹核验与评分。
