# excel to word views.py
import os
from django.http import HttpResponse
from django.conf import settings
from django.shortcuts import render
from datetime import datetime
from .forms import ExcelFileForm
from .utils import extract_tables_from_excel, create_word_document

def upload_excel(request):
    if request.method == 'POST':
        form = ExcelFileForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = form.save()
            document_name = request.POST.get('document_name', 'output_document')
            output_path = f'Extracts/{document_name}_{datetime.now().strftime("%m_%d_%Y_%H%M")}.docx'
            tables = extract_tables_from_excel(excel_file.file.path)
            create_word_document(tables, output_path)
            return render(request, 'word_processor/success.html', {'document_name': document_name, 'filename_variable': os.path.basename(output_path)})
    else:
        form = ExcelFileForm()
    return render(request, 'word_processor/upload_excel.html', {'form': form})


def download_document(request, filename):
    # Construct the file path to the document in the 'extracts' directory
    file_path = os.path.join(settings.MEDIA_ROOT, 'extracts', filename)

    # Open the document file in binary mode
    with open(file_path, 'rb') as document_file:
        # Read the contents of the document file
        document_content = document_file.read()

    # Create an HTTP response with the document content as the response body
    response = HttpResponse(document_content, content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    # Set the 'Content-Disposition' header to prompt the browser to download the file
    response['Content-Disposition'] = f'attachment; filename="{filename}"'

    return response
