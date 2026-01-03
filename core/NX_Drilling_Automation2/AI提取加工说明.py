import os
from openai import OpenAI
# "local_qwen": {
#         "base_url": "http://192.168.0.209:8000/v1",
#         "api_key": "sk-dummy",
#         "chat_model": "Qwen3-30B-A3B-Instruct",  # 与 vLLM --served-model-name 保持一致
#         "embedding_model": "text-embedding-v3",  # 使用 Qwen API 的嵌入模型
#         "embedding_base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",  # Qwen API 地址
#         "embedding_api_key": os.getenv("DASHSCOPE_API_KEY", "sk-b8c9f5f0e0e74e4e9f5e0e74e4e9f5e0"),  # Qwen API Key
#     },

prompt = """
你是一个高精度机械加工图纸解析引擎，请严格按照以下指令处理输入：

# 任务
将用户提供的“加工说明”文本逐行解析为结构化的 JSON 列表。每行以“代号 : 描述”形式存在（如 L :2 -<O>10.00割...），需提取数量、孔类型、直径、深度、螺纹、公差等信息。

# 输入处理规则（由用户提供，此处为占位说明）
- 【num】：取“:”后第一个整数，若无则默认为0。
- 【circle_type】：取“:”前的字母或字母和数字的组合。
- 【is_through_hole】：若含“钻穿”“攻穿”“铣通”“割通”等词，则 is_through_hole=true，否则is_through_hole=false。
- 【real_diamater】：若不包含“沉头”则直接从“<O>xx.x”中提取数值。若包含沉头，需要取和沉头相距较远的<O>xx.x。
- 【real_diamater_head】：若包含“沉头”，取和沉头紧跟着的<O>xx.x。若不包含沉头，默认为None。
- 【real_diamater_threading_hole】：默认为None。
- 【diamater】：默认为None。
- 【pitch】：默认为None。
- 【direction】：默认为None。
- 【luowen_depth】：默认为None。
- 【head_depth】：若包含“沉头”，取和“深”紧跟着的数，如：“L  :2 -<O>10.00割,单+0.005,背沉头<O>12.0深20.0mm(合销)”取20.0。若不包含沉头，默认为None。
- 【hole_depth】：若包含“沉头”，取和<O>xx.x紧跟着的深后的数字，如：“L  :2 -<O>10.00深35mm,背沉头<O>12.0深20.0mm(合销)”取35。若不包含沉头但包含“钻穿”“攻穿”“铣通”“割通”等词，取三个相邻数字的最后一个数字，如：“720.0L*810.0W*96.00T” 取96.00。
- 【is_bz】：根据加工说明中孔的特性判断是否背钻。判断依据：可根据出现的背沉头数量和正沉头数量的多少判断，背沉头 多 取：true，正沉头 多 取：false。若整个加工说明中无”背钻“字眼，取：false。默认为false。全局一致。
- 【main_hole_processing】：默认为钻。若包含”割、铰、铣、钻“优先级：割 > 铰 > 铣 > 钻，按顺序取。
- 【real_diamater_threading_hole】：默认为None。
- 【通孔判断】：若含“钻穿”“攻穿”“铣通”“割通”等词，则 is_through_hole=true；否则根据是否标注深度判断。
- 【沉头孔】：匹配“背沉头<O>xx.x深yy.mm”或“沉头<O>xx.x深yy.mm”。
- 【螺纹】：识别“Mxx”或“Mxx x Pyy”格式，方向由“正攻/背攻”决定。
- 【穿线孔】：提取“穿线孔<O>5.0”中的直径。
- 【加工方式】：优先级：割 > 铰 > 铣 > 钻。
- 【多行合并】：若某代号描述跨行（如 C6A 分两行），自动合并为完整描述后再解析。
- 【忽略内容】：注释（其他未注解...）等非加工行和尺寸。

# 输出格式规范
- 输出必须是 **合法 JSON 数组**，每个元素为一个对象。
- 对象字段必须完全遵循以下结构（字段缺失时设为 null，布尔值不得省略）：
{
  "num": 1,
  "circle_type": "L",
  "is_through_hole": false,
  "real_diamater": 10.0,
  "real_diamater_head": 12.0,
  "real_diamater_threading_hole": null,
  "luowen": {
    "diamater": null,
    "pitch": null,
    "direction": null,
    "luowen_depth": null
  },
  "depth": {
    "head_depth": 20.0,
    "hole_depth": null
  },
  "is_bz": true,
  "main_hole_processing": "割"
}

# 重要约束
1. 禁止添加任何解释、注释、Markdown 或额外文本。
2. 所有数值保留原始小数位数（如 10.00 → 10.0 可接受，但不要转为整数 10）。
3. 若字段无法确定，必须设为 null（JSON null），不可留空或写“未知”。
4. 确保 JSON 可被标准解析器（如 Python json.loads）直接加载。
现在，请解析以下加工说明：
"""

def build_prompt(role, prompt, task, examples=None):
    if examples:
        prompt += "示例：\n" + "\n".join(examples) + "\n\n"
    prompt += "请直接输出结果，不要解释。"
    return prompt

# # 使用
# system_msg = build_prompt(
#     role="UG工程师",
#     task="",
#     constraints="- 回答必须引用具体法条编号\n- 若不确定，请说“建议咨询执业律师”",
#     examples=[
#         "用户：租房合同没到期房东要收房怎么办？\n回答：根据《民法典》第七百零四条，租赁期限内房东不得擅自收回房屋……"
#     ]
# )

detail = """
L
L
L1
L1
L1
L1
L2
L2
L2
L2
L3
L3
M
M
M
M
M
M
M1
M1
M1
M1
M2
M2
M2
M2
M2
M2
M3
M3
M4
M4
M4
M4
M5
M5
M5
M5
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
M6
X
X
X
X
X1
X1
X1
X1
X1
X1
X2
X2
X2
C
C
C
C
C
C
C
C
C
C
C
C
C2
C2
C2
C2
C2
C2
C3
C3
C3
C4
C4
C4
C4
C4
C4
C5
C5
C6
C6
C6
C6
C6A
C6A
C6
B
B
B
B1
D
D
D
D1
D2
D2
D2
D2
D2
D2
D3
D3
D3
D4
D4
D4
D5
D6
D6
W1
W1
W1
W1
W1
W1
W1
W2
W2
W2
W2
W2
W3
W3
W3
W3
W3
W3
W3
加工说明:(下模板)_DIE-03
L  :2 -<O>10.00割,单+0.005
,背沉头<O>12.0深20.0mm(合销)
L1 :4 -<O>13.00割,单+0.005(合销)
L2 :4 -<O>16.00割,单+0.005(合销)，穿线孔<O>5.0钻
L3 :2 -<O>10.00铰深18mm(CH销)
M  :6 -<O>8.5正钻,正攻M10xP1.5深30.0mm(螺丝)
M1 :4 -<O>10.5正钻,正攻M12xP1.75深30.0mm(螺丝)
M2 :10 -10-<O>14.0钻穿,背攻M16xP2.0攻穿(螺丝)
M3 :2 -<O>14.0钻穿,背攻M16xP2.0攻穿(螺丝)
M4 :4 -<O>6.8背钻深15.0mm,
,背攻M8xP1.25深15.0mm(止付螺丝) MSW8
M5 :4 -M5,<O>6.0钻穿,(螺丝)
M6 :46-M10,<O>11.5钻穿,
,背沉头<O>18.0深25.5mm(螺丝)
X  :4 -<O>27.0钻穿(弹簧孔)
X1 :6 -<O>32.0实数铣通(弹簧孔)
X2 :3 -<O>32.0实数铣通(弹簧孔)
C  :12-<O>11.0钻穿
C2 :6 -<O>18.0钻穿
C3 :3 -<O>21.0钻穿
C4 :6 -<O>22.0钻穿
C5 :2 -<O>12.2钻穿
,背沉头<O>16.0深26.0mm
C6A :2 -<O>16.2钻穿（暂不加工，只加工背沉头）
,背沉头<O>21.0深66.0mm,深度准
B  :3 -让位,按3D实数铣
B1 :1 -让位，实数铣通
D  :3 -限位槽,正面铣深73.00mm,深度准
D1 :1 -入块槽,正面精铣深50.00mm,精铣单+0.02,深度准
D2 :6 -限位槽,正面铣深78.00mm,深度准
D3 :3 -限位槽,正面铣深64.20mm,深度准
D4 :3 -限位槽,正面铣深65.50mm,深度准
D5 :1 -限位槽,正面铣深72.20mm,深度准
D6 :2 -入块槽,正面精铣深29.32mm,精铣单+0.02,深度准
W1 :7 -入子孔,割,单+0.005，穿线孔<O>5.0钻
W2 :5 -让位，实数割通
W3 :9 -入子孔,割,单+0.005，穿线孔<O>5.0钻
720.0L*810.0W*96.00T 1PCS 45# - -
GW:439KG
其他未注解按3D涂色加工
810.00
0.00
0.00
720.00
96.00
96.00
M2
M2
M3
M3
C1
C1
M2
M2
M2
M2
M2
M2
M2
换图 2025.09.29. 赵安州
C6 :4 -<O>16.2钻穿
,背沉头<O>21.0深66.0mm,深度准
"""

try:
    client = OpenAI(
        api_key="sk-dummy",
        base_url="http://192.168.0.209:8000/v1",
    )

    completion = client.chat.completions.create(
        model="Qwen3-30B-A3B-Instruct",
        messages=[
            {'role': 'system', 'content': prompt},
            {'role': 'user', 'content': detail}
            ]
    )
    print(completion.choices[0].message.content)
except Exception as e:
    print(f"错误信息：{e}")

