from django import forms
from .models import UploadedFile
import os


class UploadFileForm(forms.ModelForm):
    class Meta:
        model = UploadedFile
        fields = ['file']

    def clean_file(self):
        file = self.cleaned_data.get('file')
        ext = os.path.splitext(file.name)[1].lower()

        allowed_extensions = ['.xlsx']
        if ext not in allowed_extensions:
            raise forms.ValidationError("Допустимы только файлы с расширением .xlsx")

        return file

