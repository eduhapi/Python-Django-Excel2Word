from django.shortcuts import render
from .forms import ExcelFileForm
from .utils import extract_tables_from_excel, create_word_document
import os

def upload_excel(request):
    if request.method == 'POST':
        form = ExcelFileForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = form.save()
            tables = extract_tables_from_excel(excel_file.file.path)
            output_path = os.path.join('Extracts', 'extracted.docx')
            create_word_document(tables, output_path)
            return render(request, 'word_processor/success.html')
    else:
        form = ExcelFileForm()
    return render(request, 'word_processor/upload_excel.html', {'form': form})
