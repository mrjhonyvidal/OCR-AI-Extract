from django import forms

class PDFUploadForm(forms.Form):
    files = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))