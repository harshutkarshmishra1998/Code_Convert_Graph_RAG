"""Import the VBA graph JSON format into Neo4j."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any

from vba_to_java.vba_graph.schema import validate_graph


NODE_QUERY = """
UNWIND $rows AS row
MERGE (n:VbaNode {id: row.id})
SET n += row.properties
"""

EDGE_QUERY = """
UNWIND $rows AS row
MATCH (source:VbaNode {id: row.source})
MATCH (target:VbaNode {id: row.target})
MERGE (source)-[r:VBA_RELATION {edge_id: row.edge_id}]->(target)
SET r += row.properties
"""


def default_graph_path(sample_number: int = 6) -> Path:
    repo_root = Path(__file__).resolve().parents[1]
    return (
        repo_root
        / "files"
        / "vba_graph_output"
        / f"sample_{sample_number}"
        / "graph.json"
    )


def _read_graph(path: Path) -> dict[str, Any]:
    graph = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(graph, dict):
        raise ValueError(f"{path}: expected a JSON object")

    errors = validate_graph(graph)
    if errors:
        raise ValueError(f"{path}: invalid graph: {'; '.join(errors)}")
    if not isinstance(graph["nodes"], list) or not isinstance(graph["edges"], list):
        raise ValueError(f"{path}: graph nodes and edges must be arrays")

    node_ids = {node.get("id") for node in graph["nodes"]}
    if any(not isinstance(node_id, str) or not node_id for node_id in node_ids):
        raise ValueError(f"{path}: every graph node must have a non-empty string ID")
    for edge in graph["edges"]:
        if not isinstance(edge.get("relation"), str):
            raise ValueError(f"{path}: every edge must have a string relation")
    return graph


def _node_row(node: dict[str, Any], project_name: str) -> dict[str, Any]:
    source = node.get("source") or {}
    properties = {
        "node_type": node.get("type", "UNKNOWN"),
        "name": node.get("name"),
        "project_name": project_name,
        "parent_id": node.get("parent_id"),
        "source_path": source.get("path"),
        "start_line": source.get("start_line"),
        "end_line": source.get("end_line"),
        "source_text": source.get("text"),
        "metadata_json": json.dumps(
            node.get("metadata") or {}, ensure_ascii=False, sort_keys=True
        ),
    }
    return {
        "id": node["id"],
        "properties": {key: value for key, value in properties.items() if value is not None},
    }


def _symbol_row(symbol: dict[str, Any], project_name: str) -> dict[str, Any]:
    properties = {
        "node_type": "SYMBOL",
        "name": symbol.get("name"),
        "project_name": project_name,
        "symbol_kind": symbol.get("kind"),
        "scope": symbol.get("scope"),
        "data_type": symbol.get("data_type"),
        "passing": symbol.get("passing"),
        "symbol_json": json.dumps(symbol, ensure_ascii=False, sort_keys=True),
    }
    return {
        "id": symbol["id"],
        "properties": {key: value for key, value in properties.items() if value is not None},
    }


def _edge_rows(graph: dict[str, Any], project_name: str) -> list[dict[str, Any]]:
    rows = []
    for edge in graph["edges"]:
        edge_data = {
            "project_name": project_name,
            "source": edge["source"],
            "target": edge["target"],
            "relation": edge["relation"],
            "evidence": edge.get("evidence") or {},
        }
        encoded = json.dumps(edge_data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        rows.append(
            {
                "edge_id": hashlib.sha256(encoded.encode("utf-8")).hexdigest(),
                "source": edge["source"],
                "target": edge["target"],
                "properties": {
                    "relation": edge["relation"],
                    "project_name": project_name,
                    "evidence_json": json.dumps(
                        edge.get("evidence") or {}, ensure_ascii=False, sort_keys=True
                    ),
                },
            }
        )
    return rows


def _prepare_import(paths: list[Path]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    node_rows = []
    edge_rows = []
    summaries = []

    for path in paths:
        graph = _read_graph(path)
        project_name = graph.get("project", {}).get("name", path.parent.name)
        graph_nodes = [_node_row(node, project_name) for node in graph["nodes"]]
        symbols = graph.get("symbols") or []
        graph_nodes.extend(_symbol_row(symbol, project_name) for symbol in symbols)
        graph_edges = _edge_rows(graph, project_name)

        node_rows.extend(graph_nodes)
        edge_rows.extend(graph_edges)
        summaries.append(
            {
                "input": str(path),
                "project": project_name,
                "nodes": len(graph["nodes"]),
                "symbols": len(symbols),
                "edges": len(graph_edges),
            }
        )

    ids = [row["id"] for row in node_rows]
    if len(ids) != len(set(ids)):
        raise ValueError("Graph inputs contain duplicate node or symbol IDs")

    return node_rows, edge_rows, summaries


def _write_graphs(
    transaction: Any,
    node_rows: list[dict[str, Any]],
    edge_rows: list[dict[str, Any]],
) -> None:
    transaction.run("MATCH (n) DETACH DELETE n").consume()
    if node_rows:
        transaction.run(NODE_QUERY, rows=node_rows).consume()
    if edge_rows:
        transaction.run(EDGE_QUERY, rows=edge_rows).consume()


def _create_driver() -> Any:
    import certifi
    import api_keys
    from neo4j import GraphDatabase

    os.environ["SSL_CERT_FILE"] = certifi.where()
    return GraphDatabase.driver(
        api_keys._loaded["NEO4J_URI"],
        auth=(
            api_keys._loaded["NEO4J_USERNAME"],
            api_keys._loaded["NEO4J_PASSWORD"],
        ),
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Import VBA graph JSON files into Neo4j."
    )
    parser.add_argument(
        "graphs",
        nargs="*",
        type=Path,
        help="Graph JSON files (default: files/vba_graph_output/sample_6/graph.json).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate inputs and report counts without connecting to Neo4j.",
    )
    args = parser.parse_args(argv)
    paths = args.graphs or [default_graph_path()]

    try:
        node_rows, edge_rows, summaries = _prepare_import(paths)
        print(json.dumps({"graphs": summaries}, indent=2, ensure_ascii=False))
        if args.dry_run:
            return 0

        driver = _create_driver()
        try:
            with driver.session() as session:
                session.execute_write(
                    _write_graphs, node_rows, edge_rows
                )
        finally:
            driver.close()
        print(f"Imported {len(node_rows)} nodes and {len(edge_rows)} relationships.")
        return 0
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    n = 6
    default_input = default_graph_path(n)
    raise SystemExit(main(sys.argv[1:] or [str(default_input)]))
