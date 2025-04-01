from django.urls import path
from .views import upload_file, file_list, process_with_chatgpt

urlpatterns = [
    path("upload/", upload_file, name="upload_file"),
    path("files/", file_list, name="file_list"),
    path("process-ai/<int:file_id>/", process_with_chatgpt, name="process_with_chatgpt"),
]
