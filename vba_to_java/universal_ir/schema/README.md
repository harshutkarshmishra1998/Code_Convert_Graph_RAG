# Universal IR schema

`ir.schema.json` is the JSON Schema Draft 2020-12 contract for IR version `1.0.0`.

The Python validator adds checks that are awkward to express in JSON Schema, including unique IDs, duplicate labels, and simple cross-reference integrity. The schema deliberately permits additional properties so source-language semantic annotations and future IR extensions can be added without discarding information. New core node kinds still require an explicit versioned schema update.
