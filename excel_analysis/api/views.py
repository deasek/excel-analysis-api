from rest_framework.views import APIView
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework.response import Response
from rest_framework import status
from drf_spectacular.utils import extend_schema
from .serializers import ExcelUploadSerializer, ExcelAnalysisResponseSerializer
from .utils import process_excel_file


class ExcelAnalysisView(APIView):
    parser_classes = [MultiPartParser, FormParser]

    @extend_schema(
        request=ExcelUploadSerializer,
        responses={200: ExcelAnalysisResponseSerializer},
        description="Upload Excel file and analyze specified columns",
        summary="Analyze Excel File",
        tags=["Excel Analysis"],
    )
    def post(self, request):
        try:
            serializer = ExcelUploadSerializer(data=request.data)

            if not serializer.is_valid():
                return Response(
                    {"error": "Invalid input", "details": serializer.errors},
                    status=status.HTTP_400_BAD_REQUEST,
                )

            file = serializer.validated_data["file"]
            columns = serializer.validated_data["columns"]
            sheet_name = serializer.validated_data.get("sheet_name")

            summary = process_excel_file(file, columns, sheet_name)

            response_data = {"file": file.name, "summary": summary}

            response_serializer = ExcelAnalysisResponseSerializer(data=response_data)
            if response_serializer.is_valid():
                return Response(
                    response_serializer.validated_data, status=status.HTTP_200_OK
                )
            else:
                return Response(response_data, status=status.HTTP_200_OK)

        except Exception as e:
            return Response(
                {"error": f"Failed to process file: {str(e)}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )
