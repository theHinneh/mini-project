import os
import warnings
from openpyxl import load_workbook
warnings.filterwarnings('ignore')
from django.contrib.staticfiles.templatetags.staticfiles import static
from django.core.files.storage import FileSystemStorage
from django.shortcuts import render
from django.views.generic import TemplateView
from zipfile import ZipFile 
import zipfile
import os
from wsws.settings import BASE_DIR
from .utils import read_file
import openpyxl

class FileView(TemplateView):
    template_name = 'index.html'

    def post(self, request, *args, **kwargs):

        files = request.FILES
        fs = FileSystemStorage()
        

        
        file_4 = static(r'\wsws_file\renamed_files.zip')
        file_4_path = BASE_DIR + file_4
        output = []
        output += read_file(files.get('lmi'), files.get('cl'), files.get('ts'), file_4_path)
        output.append(file_4 if 'static' in file_4 else fs.url(file_4))
        output = map(lambda x: {"type": x.split('.')[1], "url": request.build_absolute_uri(x)}, output)
        return render(request, self.template_name, {"files": output})
