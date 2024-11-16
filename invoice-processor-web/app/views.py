from django.shortcuts import render
from .forms import PDFUploadForm
from .utils.pdf_extractor import extract_text_from_pdf
from .utils.openai_helper import extract_data_with_openai
from .utils.webhook_sender import send_to_webhook

def dashboard(request):
    if request.method == 'POST':
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            files = request.FILES.getlist('files')
            results = []
            for file in files:
                text = extract_text_from_pdf(file, "/opt/homebrew/bin")
                extracted_data = extract_data_with_openai(text)
                success = send_to_webhook(extracted_data)
                results.append((file.name, success))
            return render(request, 'results.html', {'results': results})
    else:
        form = PDFUploadForm()
    return render(request, 'dashboard.html', {'form': form})