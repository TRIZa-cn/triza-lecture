# -*- coding: utf-8 -*-
import json, os
D = r"D:\doc\ZGCLab\triza\project\design\akbank"
os.makedirs(os.path.join(D, "schema"), exist_ok=True)
os.makedirs(os.path.join(D, "seed"), exist_ok=True)

schema = {
 "_status": "DESIGN-DRAFT-NOT-FOR-IMPORT",
 "version": "0.1-design",
 "entity": "constraint_candidate",
 "note": "CC 是 PC 的前体：parameter/ends/pcvt/separation/principles/verification 六字段与 contradiction.schema.json#/definitions/physical 的 required 同名同构，promote 为字段拷贝而非翻译。",
 "required": ["id", "type", "dimension", "exclusion", "guidance", "parameter", "ends",
              "pcvt", "separation", "principles", "owner", "revocable",
              "abandon_condition", "level_shift", "verification"],
 "properties": {
  "id": {"type": "string", "pattern": "^CC-[0-9]{3}$"},
  "type": {"const": "constraint_pair"},
  "dimension": {
   "enum": ["target", "resource", "method", "subject"],
   "note": "改哪一段：target=做什么 resource=用什么 method=怎么做 subject=核心对象"
  },
  "exclusion": {
   "type": "object",
   "required": ["object", "why_shortcut"],
   "properties": {
    "object": {"type": "string", "note": "被禁掉的那条最省力路径，必须能被指认"},
    "why_shortcut": {"type": "string", "note": "凭什么说它最省力"},
    "evidence": {"type": "string", "note": "证据来源；无证据则写 assumed"}
   }
  },
  "guidance": {
   "type": "object",
   "required": ["action", "measurable"],
   "properties": {
    "action": {"type": "string", "note": "具体可执行动作，不得是方向词"},
    "measurable": {"type": "string", "note": "怎么算做到了；空串视为不合格"}
   }
  },
  "parameter": {"type": "string", "pattern": "^(SEC-P[0-9]{2,3}|CLS-P[0-9]{2,3})$"},
  "ends": {
   "type": "array", "minItems": 2, "maxItems": 2,
   "items": {
    "type": "object",
    "required": ["value", "requirement", "rationale"],
    "properties": {
     "value": {"type": "string"}, "requirement": {"type": "string"}, "rationale": {"type": "string"}
    }
   }
  },
  "pcvt": {
   "type": "object",
   "required": ["end1_independently_valid", "end2_independently_valid", "verdict"],
   "properties": {
    "end1_independently_valid": {"type": "boolean"},
    "end2_independently_valid": {"type": "boolean"},
    "falsification_attempted": {"type": "string"},
    "verdict": {"enum": ["valid_physical", "downgrade_to_technical"]}
   }
  },
  "separation": {
   "type": "object", "required": ["kind"],
   "properties": {
    "kind": {"enum": ["time", "space", "condition", "whole_part"]},
    "axis": {"type": "string"}
   }
  },
  "principles": {"type": "array", "items": {"type": "integer"}},
  "owner": {"enum": ["self", "external"], "note": "约束由谁设定"},
  "revocable": {"type": "boolean", "note": "能否在不承担不可逆后果的前提下撤掉"},
  "abandon_condition": {"type": "string", "note": "什么条件下必须放弃该约束；空串视为不合格"},
  "level_shift": {
   "type": "object", "required": ["from", "to"],
   "properties": {"from": {"type": "integer"}, "to": {"type": "integer"}}
  },
  "verification": {
   "type": "object", "required": ["metrics"],
   "properties": {"contract_id": {"type": "string"}, "metrics": {"type": "array"}}
  },
  "promoted_to": {"type": ["string", "null"], "note": "落库后的 PC-xx；未落库为 null"},
  "sources": {"type": "array"}
 },
 "governance": {
  "R1": "exclusion.object 必须来自输入中的 current_solution，不得凭空发明最省力路径",
  "R2": "guidance.measurable 为空串即不合格：只给方向词不算引导",
  "R3": "两条必须落在同一对象上；指认不出同一对象即降级为技术矛盾（pcvt.verdict=downgrade_to_technical）",
  "R4": "abandon_condition 为空串即不合格：防止过早收敛（premature closure）",
  "no_scoring": "本 schema 不含打分字段；CC 只保证形状合规，不判断约束好坏"
 }
}
json.dump(schema, open(os.path.join(D, "schema", "constraint_candidate.schema.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

seed = {
 "_status": "DESIGN-DRAFT-NOT-FOR-IMPORT",
 "entity": "constraint_candidate",
 "count": 2,
 "items": [
  {
   "id": "CC-001",
   "type": "constraint_pair",
   "dimension": "method",
   "exclusion": {
    "object": "按角色打包授权：read / write / delete 一次性授予，运营侧不为每个操作类型单独建模",
    "why_shortcut": "一次性授权免去逐操作枚举与维护，落地成本最低",
    "evidence": "assumed（来自 PC-01 讨论中的现有实现假设，待实测确认）"
   },
   "guidance": {
    "action": "要求最小可授予单元 = 单一操作类型，运营侧必须提交操作清单，未列入清单的操作不可授予",
    "measurable": "权限矩阵中可独立授予的操作维度数 ≤ 1"
   },
   "parameter": "SEC-P03",
   "ends": [
    {"value": "粗", "requirement": "授权粒度必须粗",
     "rationale": "粒度过细则智能体无法获得调用权限，自主性归零"},
    {"value": "细", "requirement": "授权粒度必须细",
     "rationale": "粒度过粗则须显式授予 read，凭据暴露面扩大"}
   ],
   "pcvt": {
    "end1_independently_valid": True,
    "end2_independently_valid": True,
    "falsification_attempted": "尝试用授权代理否证『必须有 read 权限才能完成调用』：中介物代持凭据并校验请求，调用可达而明文不可读，故该前提被否证，但两端各自仍成立",
    "verdict": "valid_physical"
   },
   "separation": {"kind": "condition", "axis": "操作类型"},
   "principles": [24],
   "owner": "self",
   "revocable": True,
   "abandon_condition": "若授权代理引入的端到端额外时延 > 200 ms 且无法用本地缓存消除，则放弃该约束，回退到按风险分级的粗粒度授权",
   "level_shift": {"from": 2, "to": 3},
   "verification": {"contract_id": "MC-P03", "metrics": ["越权拦截率", "凭据零残留", "TTL 合规率"]},
   "promoted_to": "PC-01",
   "sources": ["TRIZa_网安参数与矛盾库_设计稿.md §A.1", "§D PC-01"]
  },
  {
   "id": "CC-002",
   "type": "constraint_pair",
   "dimension": "method",
   "exclusion": {
    "object": "所有受控操作一律走事前人工审批，不区分操作是否可撤回",
    "why_shortcut": "不区分即无需为每个操作建模影响范围与补偿动作，实现最省力",
    "evidence": "assumed（审批流程现状，待实测统计人工介入率）"
   },
   "guidance": {
    "action": "每个受控操作必须声明补偿动作与撤回时限；未声明的操作不得上线，可逆操作默认放行并事后审计",
    "measurable": "受控操作中已声明补偿动作的比例 = 100%；未补偿操作数 = 0"
   },
   "parameter": "SEC-P13",
   "ends": [
    {"value": "高", "requirement": "自主度必须高",
     "rationale": "可逆操作也送审会把人变成瓶颈，审批疲劳导致盖章式通过"},
    {"value": "低", "requirement": "自主度必须低",
     "rationale": "不可逆操作（转账、外发、删库、对外发布）一旦放行无法收回"}
   ],
   "pcvt": {
    "end1_independently_valid": True,
    "end2_independently_valid": True,
    "falsification_attempted": "尝试用『全部操作都做成可撤销』否证低自主度的必要性：外部副作用（已发出的邮件、已公开的发布、第三方状态变更）无法撤回，故否证失败，两端均成立",
    "verdict": "valid_physical"
   },
   "separation": {"kind": "condition", "axis": "操作可逆性（SEC-P25）"},
   "principles": [10, 34],
   "owner": "self",
   "revocable": True,
   "abandon_condition": "若统计显示不可逆操作占比 > 30%，说明系统本身缺乏可逆设计，应先做架构改造，此约束暂缓",
   "level_shift": {"from": 3, "to": 4},
   "verification": {"contract_id": "MC-P13", "metrics": ["人工介入率", "未补偿操作数", "撤回成功率"]},
   "promoted_to": None,
   "sources": ["A 讲第一讲三层写法讨论（2026-10-29）", "SEC-P25 操作可逆性（本轮新增）"]
  }
 ],
 "rejected_examples": [
  {"guidance": "让它更安全一些", "reason": "违反 R2：measurable 为空，只给方向词不算引导"},
  {"exclusion": "不要用旧方法", "reason": "违反 R1：exclusion.object 未指向具体的最省力路径"},
  {"abandon_condition": "", "reason": "违反 R4：缺少放弃条件，构成过早收敛风险"}
 ]
}
json.dump(seed, open(os.path.join(D, "seed", "constraint_pairs.seed.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("schema/seed written")
for f in ["schema/constraint_candidate.schema.json", "seed/constraint_pairs.seed.json"]:
    p = os.path.join(D, f); print(" ", f, round(os.path.getsize(p)/1024, 1), "KB")