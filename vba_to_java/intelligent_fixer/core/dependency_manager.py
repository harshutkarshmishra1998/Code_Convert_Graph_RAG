"""Artifact dependency graph and safe rebuild planning."""
from collections import defaultdict, deque

DEPENDENCIES = {
 "graph": ["source"],
 "segments": ["source","graph"],
 "ir": ["source","graph","segments"],
 "normalized_ir": ["ir","source"],
 "semantic_analysis": ["normalized_ir","graph","segments","source"],
 "pseudocode": ["graph","segments","source"],
 "pseudocode_run": ["pseudocode","graph"],
 "report": ["graph","segments"],
}
class DependencyManager:
    def __init__(self, dependencies=None):
        self.dependencies=dependencies or DEPENDENCIES
        self.children=defaultdict(list)
        for child, parents in self.dependencies.items():
            for parent in parents: self.children[parent].append(child)

    def downstream(self, changed):
        seen=set(); q=deque(changed)
        while q:
            x=q.popleft()
            for child in self.children.get(x,[]):
                if child not in seen: seen.add(child); q.append(child)
        return sorted(seen)

    def rebuild_order(self, changed):
        affected=set(self.downstream(changed))
        order=[]; remaining=set(affected)
        while remaining:
            ready=sorted(x for x in remaining if not (set(self.dependencies.get(x,[])) & remaining))
            if not ready: raise ValueError("Dependency cycle detected")
            order.extend(ready); remaining.difference_update(ready)
        return order

    def plan(self, changed_artifacts):
        return {"changed":sorted(set(changed_artifacts)),
                "invalidate_or_rebuild":self.rebuild_order(changed_artifacts),
                "note":"Downstream artifacts are stale until rebuilt and revalidated."}
