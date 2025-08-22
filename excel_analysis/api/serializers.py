from rest_framework import serializers


class ExcelUploadSerializer(serializers.Serializer):
    file = serializers.FileField(help_text="Excel file to upload (.xlsx or .xls)")
    columns = serializers.ListField(
        child=serializers.CharField(),
        help_text="List of column names to analyze",
        allow_empty=False,
    )
    sheet_name = serializers.CharField(
        help_text="Optional sheet name to analyze",
        required=False,
        allow_blank=True,
        default=None,
    )

    def validate_file(self, value):
        if not value.name.lower().endswith((".xlsx", ".xls")):
            raise serializers.ValidationError(
                "File must be an Excel file (.xlsx or .xls)"
            )
        return value

    def validate_columns(self, value):
        if not value:
            raise serializers.ValidationError(
                "At least one column name must be provided"
            )

        for column in value:
            if not isinstance(column, str) or not column.strip():
                raise serializers.ValidationError(
                    "All column names must be non-empty strings"
                )

        return [col.strip() for col in value]


class ColumnSummarySerializer(serializers.Serializer):
    column = serializers.CharField(help_text="Column name")
    sum = serializers.FloatField(help_text="Sum of all numeric values")
    avg = serializers.FloatField(help_text="Average of all numeric values")


class ExcelAnalysisResponseSerializer(serializers.Serializer):
    file = serializers.CharField(help_text="Original filename")
    summary = ColumnSummarySerializer(many=True)
