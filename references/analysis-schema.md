# Review JSON

The deterministic pipeline creates all nodes, hierarchy edges, citation edges, unresolved references, and statistics. The Agent may only submit corrections that are visible in supplied material.

```json
{
  "source_identity_overrides": [
    {
      "source_id": "SRC-00001",
      "title": "用户确认的法规名称",
      "reason": "文件名是扫描件编号，正文标题清晰可见"
    }
  ],
  "review_items": [
    {
      "source_id": "SRC-00001",
      "node_id": "SRC-00001:A-1",
      "field": "编号",
      "reason": "扫描件条号模糊，需人工回看原件"
    }
  ]
}
```

`source_id` and `node_id` must exist in the bundle. Overrides change display labels only; they never alter hierarchy, citations, source excerpts, or parsed identifiers. Omit both arrays when no review is needed.
