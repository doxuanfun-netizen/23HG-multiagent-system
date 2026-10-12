# -*- coding: utf-8 -*-
"""
Trình tải cấu hình miền (Domain Router) từ file YAML.

Tự động nạp bộ định tuyến và ánh xạ lệnh từ file cấu hình YAML,
hỗ trợ mở rộng định tuyến lệnh bằng ngôn ngữ tự nhiên.
"""

import ast
import logging
import os
from typing import Any, Dict, List, Optional

from .base_router import BaseCommandRouter
from .types import CommandCandidate
from .scorer import ScoringEngine

logger = logging.getLogger('domain_loader')


def _safe_eval_ast(node: ast.AST, variables: Dict[str, Any]) -> Any:
    """Đánh giá an toàn cây cú pháp trừu tượng (AST) không sử dụng eval()."""
    if isinstance(node, ast.Expression):
        return _safe_eval_ast(node.body, variables)
    elif isinstance(node, ast.Constant):
        return node.value
    elif isinstance(node, ast.Name):
        if node.id in variables:
            return variables[node.id]
        raise ValueError(f"Biến không được phép: {node.id}")
    elif isinstance(node, ast.Compare):
        left = _safe_eval_ast(node.left, variables)
        for op, comparator in zip(node.ops, node.comparators):
            right = _safe_eval_ast(comparator, variables)
            if isinstance(op, ast.In):
                if left not in right:
                    return False
            elif isinstance(op, ast.NotIn):
                if left in right:
                    return False
            elif isinstance(op, ast.Eq):
                if left != right:
                    return False
            elif isinstance(op, ast.NotEq):
                if left == right:
                    return False
            else:
                raise ValueError(f"Toán tử so sánh không được hỗ trợ: {type(op)}")
        return True
    elif isinstance(node, ast.BoolOp):
        if isinstance(node.op, ast.And):
            return all(_safe_eval_ast(val, variables) for val in node.values)
        elif isinstance(node.op, ast.Or):
            return any(_safe_eval_ast(val, variables) for val in node.values)
    elif isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
        return not _safe_eval_ast(node.operand, variables)
    raise ValueError(f"Cú pháp biểu thức không được phép: {type(node)}")


def _validate_ast_nodes(node: ast.AST):
    ALLOWED_NODE_TYPES = (
        ast.Expression, ast.Constant, ast.Name, ast.Compare, ast.BoolOp,
        ast.UnaryOp, ast.In, ast.NotIn, ast.Eq, ast.NotEq, ast.And, ast.Or, ast.Not,
        ast.Load
    )
    for n in ast.walk(node):
        if not isinstance(n, ALLOWED_NODE_TYPES):
            raise ValueError(f"Cú pháp hoặc nút AST nguy hiểm không được phép: {type(n).__name__}")

def _compile_trigger(expr: str):
    """
    Biên dịch biểu thức điều kiện an toàn qua AST thay vì eval().
    Chỉ cho phép các toán tử logic an toàn: 'in', 'not in', 'and', 'or', '==', '!='.
    """
    parsed = ast.parse(expr.strip(), mode='eval')
    _validate_ast_nodes(parsed)
    return lambda intent: _safe_eval_ast(parsed, {"intent": intent})


class YamlDomainRouter(BaseCommandRouter):
    """Bộ định tuyến miền được khởi tạo từ cấu hình YAML."""

    def __init__(self, config: Dict[str, Any]):
        self._config = config
        self._domain = config["domain"]
        self._intents = config.get("intents", {})
        self._entities = config.get("entities", {})
        self._commands = config.get("commands", {})
        self._negative_rules_config = config.get("negative_rules", [])
        self._disambiguation_hints = config.get("disambiguation_hints", [])
        super().__init__()

    @property
    def domain_name(self) -> str:
        return self._domain

    def _load_taxonomy(self):
        """Dữ liệu phân loại đã được cấu hình trong __init__."""
        pass

    def _get_intent_categories(self) -> Dict:
        return self._intents

    def _get_command_mappings(self) -> Dict:
        return self._commands

    def _get_negative_rules(self) -> List:
        rules = []
        for rule_cfg in self._negative_rules_config:
            trigger_expr = rule_cfg.get("trigger", "")
            wrong = rule_cfg.get("wrong_commands", [])
            penalty = rule_cfg.get("penalty", 0.3)
            reason = rule_cfg.get("reason", "")

            try:
                trigger_fn = _compile_trigger(trigger_expr)
            except Exception as e:
                logger.warning(f"Không thể biên dịch điều kiện kích hoạt quy tắc: {trigger_expr} — {e}")
                continue

            rules.append({
                "trigger": trigger_fn,
                "wrong": wrong,
                "penalty": penalty,
                "why_not": reason,
            })
        return rules

    def _recognize_entities(self, user_intent: str) -> List[str]:
        recognized = []
        for entity_name, entity_def in self._entities.items():
            keywords = entity_def.get("keywords", [])
            if any(kw in user_intent for kw in keywords):
                recognized.append(entity_name)
        return recognized

    def _generate_candidates(
        self,
        intent: Dict,
        entities: List[str],
        user_intent: str,
        context: Optional[Dict]
    ) -> List[CommandCandidate]:
        candidates = []
        for cmd_name, cmd_def in self._commands.items():
            score = self.scorer.calculate_score(
                intent_match=(intent["category"] == cmd_def.get("intent", "")),
                entity_match=bool(set(entities) & set(cmd_def.get("entities", []))),
                keyword_matches=[kw for kw in cmd_def.get("keywords", []) if kw in user_intent],
                anti_pattern_matches=[kw for kw in cmd_def.get("anti_keywords", []) if kw in user_intent],
                differentiator_match=any(d in user_intent for d in cmd_def.get("differentiators", [])),
                context=context,
                context_features=cmd_def.get("context_features", {})
            )

            full_cmd_name = f"api_{self._domain}_{cmd_name}"
            candidates.append(CommandCandidate(
                command=full_cmd_name,
                score=score,
                category=cmd_def.get("intent", "unknown"),
                reason=cmd_def.get("description", ""),
                why_not=None,
                usage_example=cmd_def.get("usage_example")
            ))
        return candidates

    def _generate_disambiguation_hint(
        self,
        top_two: List[CommandCandidate],
        user_intent: str
    ) -> Optional[str]:
        if len(top_two) < 2:
            return None

        score_diff = abs(top_two[0].score - top_two[1].score)
        if score_diff >= 0.3:
            return None

        cmd_pair = {top_two[0].command, top_two[1].command}
        for hint in self._disambiguation_hints:
            hint_pair = set(hint.get("pair", []))
            hint_pair_full = {f"api_{self._domain}_{c}" for c in hint_pair}
            if cmd_pair & hint_pair_full:
                return hint.get("hint", "")

        return f"2 lệnh có điểm số gần nhau: {top_two[0].command} ({top_two[0].score:.2f}) vs {top_two[1].command} ({top_two[1].score:.2f})"


def load_domain_from_yaml(yaml_path: str) -> YamlDomainRouter:
    """Tải bộ định tuyến miền từ file cấu hình YAML."""
    try:
        import yaml
    except ImportError:
        raise ImportError("Cần cài đặt thư viện PyYAML: pip install pyyaml")

    with open(yaml_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    if "domain" not in config:
        raise ValueError(f"YAML thiếu trường bắt buộc 'domain': {yaml_path}")

    logger.info(f"Đã tải bộ định tuyến miền từ YAML: {config['domain']} ({yaml_path})")
    return YamlDomainRouter(config)


def discover_yaml_domains(domains_dir: str) -> Dict[str, YamlDomainRouter]:
    """Tự động phát hiện và nạp tất cả cấu hình domain.yaml trong thư mục domains."""
    routers = {}

    if not os.path.isdir(domains_dir):
        return routers

    for entry in os.listdir(domains_dir):
        yaml_path = os.path.join(domains_dir, entry, "domain.yaml")
        if os.path.isfile(yaml_path):
            try:
                router = load_domain_from_yaml(yaml_path)
                routers[router.domain_name] = router
                logger.info(f"Phát hiện miền YAML: {router.domain_name}")
            except Exception as e:
                logger.warning(f"Không thể nạp cấu hình miền {yaml_path}: {e}")

    return routers
