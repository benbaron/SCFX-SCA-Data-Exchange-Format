#!/usr/bin/env python3
"""SCLX JSON validator.

Loads one or more JSON Schema files at startup, selects a schema, then validates
an input JSON document and prints a structured report.

Typical usage:
    python sclx_json_validator.py --schema-dir /path/to/schemas --input file.json
    python sclx_json_validator.py --schema /path/to/sclx-1.2-full.schema.json --input file.json

Selection behavior:
- If --schema is provided, that schema is used.
- Else, if the input JSON has a top-level "version", the tool looks for a file
  named like "sclx-<version>-full.schema.json" in --schema-dir.
- Else, if exactly one schema file is present in --schema-dir, that schema is used.

Exit codes:
    0 = valid
    1 = invalid JSON against schema
    2 = usage/runtime/schema loading error
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from jsonschema import Draft202012Validator, exceptions
from jsonschema.validators import validator_for
from referencing import Registry, Resource


SCHEMA_SUFFIX = ".schema.json"


@dataclass(frozen=True)
class LoadedSchema:
    """A loaded JSON Schema document plus file metadata."""

    path: Path
    content: Dict[str, Any]
    schema_id: Optional[str]
    title: Optional[str]
    version_const: Optional[str]


@dataclass(frozen=True)
class ValidationIssue:
    """One validation error rendered in a stable, user-friendly shape."""

    instance_path: str
    schema_path: str
    validator: str
    message: str
    value_snippet: str


@dataclass(frozen=True)
class ValidationReport:
    """Overall validation report."""

    input_path: Path
    schema_path: Path
    schema_title: Optional[str]
    schema_id: Optional[str]
    document_version: Optional[str]
    valid: bool
    issues: Tuple[ValidationIssue, ...]


class SchemaCatalog:
    """Loads schema files up front and exposes schema-selection helpers."""

    def __init__(self, schemas: Sequence[LoadedSchema], registry: Registry):
        self.schemas = list(schemas)
        self.registry = registry
        self._by_path = {schema.path.resolve(): schema for schema in self.schemas}
        self._by_filename = {schema.path.name: schema for schema in self.schemas}
        self._by_version = {
            schema.version_const: schema
            for schema in self.schemas
            if schema.version_const
        }

    @classmethod
    def from_paths(cls, schema_paths: Iterable[Path]) -> "SchemaCatalog":
        schemas: List[LoadedSchema] = []
        registry = Registry()

        for path in sorted({p.resolve() for p in schema_paths}):
            with path.open("r", encoding="utf-8") as handle:
                content = json.load(handle)

            # Fail fast on malformed schemas.
            Draft202012Validator.check_schema(content)

            schema = LoadedSchema(
                path=path,
                content=content,
                schema_id=_string_or_none(content.get("$id")),
                title=_string_or_none(content.get("title")),
                version_const=_extract_version_const(content),
            )
            schemas.append(schema)

            # Register the schema under both its $id (if present) and file URI.
            resource = Resource.from_contents(content)
            if schema.schema_id:
                registry = registry.with_resource(schema.schema_id, resource)
            registry = registry.with_resource(path.as_uri(), resource)

        if not schemas:
            raise ValueError("No schema files were loaded.")

        return cls(schemas, registry)

    def pick_schema(
        self,
        *,
        explicit_schema: Optional[Path],
        document: Dict[str, Any],
    ) -> LoadedSchema:
        """Select the schema to validate against."""

        if explicit_schema is not None:
            resolved = explicit_schema.resolve()
            try:
                return self._by_path[resolved]
            except KeyError as exc:
                raise ValueError(f"Schema not loaded: {resolved}") from exc

        document_version = _string_or_none(document.get("version"))
        if document_version:
            expected_name = f"sclx-{document_version}-full.schema.json"
            by_name = self._by_filename.get(expected_name)
            if by_name is not None:
                return by_name

            by_version = self._by_version.get(document_version)
            if by_version is not None:
                return by_version

        if len(self.schemas) == 1:
            return self.schemas[0]

        available = ", ".join(schema.path.name for schema in self.schemas)
        raise ValueError(
            "Could not determine which schema to use. "
            "Pass --schema explicitly or validate a document with a top-level "
            f'"version" that matches one of: {available}'
        )


def _string_or_none(value: Any) -> Optional[str]:
    return value if isinstance(value, str) and value.strip() else None


def _extract_version_const(schema: Dict[str, Any]) -> Optional[str]:
    """Try to extract a top-level version const from the schema."""

    try:
        version = schema["properties"]["version"]["const"]
    except Exception:
        return None
    return version if isinstance(version, str) else None


def discover_schema_files(schema_dir: Path) -> List[Path]:
    """Return all candidate schema files under one directory."""

    return sorted(
        path
        for path in schema_dir.iterdir()
        if path.is_file() and path.name.endswith(SCHEMA_SUFFIX)
    )


def load_json_file(path: Path) -> Dict[str, Any]:
    """Load a JSON object from disk and enforce top-level object shape."""

    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)

    if not isinstance(data, dict):
        raise ValueError("Input JSON must be a top-level object.")
    return data


def build_validator(schema: LoadedSchema, registry: Registry):
    """Instantiate the appropriate validator class for one loaded schema."""

    validator_cls = validator_for(schema.content)
    validator_cls.check_schema(schema.content)
    return validator_cls(schema.content, registry=registry)


def validate_document(
    input_path: Path,
    document: Dict[str, Any],
    schema: LoadedSchema,
    registry: Registry,
) -> ValidationReport:
    """Validate one document and return a structured report."""

    validator = build_validator(schema, registry)
    errors = sorted(validator.iter_errors(document), key=_error_sort_key)

    issues = tuple(
        ValidationIssue(
            instance_path=_format_path(error.absolute_path),
            schema_path=_format_path(error.absolute_schema_path),
            validator=error.validator,
            message=error.message,
            value_snippet=_safe_value_snippet(error.instance),
        )
        for error in errors
    )

    return ValidationReport(
        input_path=input_path,
        schema_path=schema.path,
        schema_title=schema.title,
        schema_id=schema.schema_id,
        document_version=_string_or_none(document.get("version")),
        valid=not issues,
        issues=issues,
    )


def _error_sort_key(error: exceptions.ValidationError) -> Tuple[str, str, str]:
    return (
        _format_path(error.absolute_path),
        _format_path(error.absolute_schema_path),
        error.message,
    )


def _format_path(path_parts: Iterable[Any]) -> str:
    parts = list(path_parts)
    if not parts:
        return "$"

    rendered = "$"
    for part in parts:
        if isinstance(part, int):
            rendered += f"[{part}]"
        else:
            rendered += f".{part}"
    return rendered


def _safe_value_snippet(value: Any, max_len: int = 160) -> str:
    try:
        text = json.dumps(value, ensure_ascii=False, default=str)
    except TypeError:
        text = repr(value)
    if len(text) > max_len:
        return text[: max_len - 3] + "..."
    return text


def report_as_text(report: ValidationReport) -> str:
    """Render a human-friendly validation report."""

    lines = [
        f"Input file:      {report.input_path}",
        f"Schema file:     {report.schema_path}",
        f"Schema title:    {report.schema_title or '(untitled)'}",
        f"Schema id:       {report.schema_id or '(none)'}",
        f"Document version:{' ' + report.document_version if report.document_version else ' (missing)'}",
        f"Valid:           {'YES' if report.valid else 'NO'}",
        f"Issue count:     {len(report.issues)}",
    ]

    if report.issues:
        lines.append("")
        lines.append("Validation issues:")
        for index, issue in enumerate(report.issues, start=1):
            lines.extend(
                [
                    f"  {index}. {issue.message}",
                    f"     Instance path: {issue.instance_path}",
                    f"     Schema path:   {issue.schema_path}",
                    f"     Validator:     {issue.validator}",
                    f"     Value:         {issue.value_snippet}",
                ]
            )
    return "\n".join(lines)


def report_as_json(report: ValidationReport) -> str:
    """Render a machine-friendly JSON report."""

    payload = {
        "inputFile": str(report.input_path),
        "schemaFile": str(report.schema_path),
        "schemaTitle": report.schema_title,
        "schemaId": report.schema_id,
        "documentVersion": report.document_version,
        "valid": report.valid,
        "issueCount": len(report.issues),
        "issues": [
            {
                "instancePath": issue.instance_path,
                "schemaPath": issue.schema_path,
                "validator": issue.validator,
                "message": issue.message,
                "valueSnippet": issue.value_snippet,
            }
            for issue in report.issues
        ],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate an SCLX JSON file against a JSON Schema. "
            "Loads schema files up front, then validates one input file."
        )
    )
    parser.add_argument(
        "--schema-dir",
        type=Path,
        default=Path.cwd(),
        help="Directory containing one or more *.schema.json files.",
    )
    parser.add_argument(
        "--schema",
        type=Path,
        help="Specific schema file to use. Must also be inside --schema-dir or be loaded explicitly.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="JSON file to validate.",
    )
    parser.add_argument(
        "--report-format",
        choices=("text", "json"),
        default="text",
        help="Output report format.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    try:
        schema_paths = discover_schema_files(args.schema_dir)
        if args.schema is not None and args.schema.resolve() not in {p.resolve() for p in schema_paths}:
            schema_paths.append(args.schema)

        catalog = SchemaCatalog.from_paths(schema_paths)
        document = load_json_file(args.input)
        schema = catalog.pick_schema(explicit_schema=args.schema, document=document)
        report = validate_document(args.input, document, schema, catalog.registry)
    except json.JSONDecodeError as exc:
        print(
            f"JSON parse error in {args.input}: line {exc.lineno}, column {exc.colno}: {exc.msg}",
            file=sys.stderr,
        )
        return 2
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if args.report_format == "json":
        print(report_as_json(report))
    else:
        print(report_as_text(report))

    return 0 if report.valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
