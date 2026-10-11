# 一次生成的数据格式
JSON只来自最后确认的正文，不让用户填写JSON。姓名和目标岗位可空；personal为需要保留的字段，缺性别/年龄/手机/邮箱会标待补充。已确认删除的籍贯等不传入。photo是实际照片路径，不传则使用待上传标识，不能填模板照片。
summary_original保留原有有效个人介绍逐条文字；summary_additions只填本轮新增、用户确认的核心亮点；自动去重但不替换原文。新建且无原介绍时original可空，已有介绍不能为方便生成而清空。用户明确不展示才设置omit_summary:true及summary_omission_confirmed:true。
sections按确认顺序传入，每个含kind/title/entries；kind为education/work/internship/project/campus/skills/awards。每条entry支持heading/period/detail/bullets。无内容模块entries为空会自动移除。独立实践等自定义模块须实际独立项目且用户同意，填写independent_section_confirmed:true；不可用此字段绕过归位规则。

```json
{
 "confirmed":true,"template_id":"T01","name":"示例候选人","target":"短视频编导",
 "personal":{"性别":"待补充","年龄":"待补充","手机":"待补充","邮箱":"待补充"},
 "summary_original":["具备品牌内容策划与多平台运营经验。"],
 "summary_additions":["能够使用剪映完成视频剪辑，并有室内拍摄实践。"],
 "sections":[
  {"kind":"education","title":"教育背景","entries":[{"heading":"示例大学 | 市场营销 | 本科","period":"2023.09—2027.06"}]},
  {"kind":"internship","title":"实习经历","entries":[{"heading":"示例公司 | 市场实习生","period":"2025.06—2025.12","bullets":["参与新品发布内容制作与账号日常运营。"]}]},
  {"kind":"awards","title":"荣誉奖项","entries":[]}
 ]
}
```
这是合成格式示例，不能把其中的姓名、日期、经历复制给真实用户。照片相对路径相对于JSON所在目录解析。输出不得覆盖源文件或已有成品。首轮和修正使用不同输出文件名，共用同一个session.json。
