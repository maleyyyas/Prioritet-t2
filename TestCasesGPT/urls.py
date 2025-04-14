import os

from django.conf import settings
from django.conf.urls.static import static
from django.urls import path

from processor.views import upload_and_list_files, process_with_chatgpt

urlpatterns = [
    path("", upload_and_list_files, name="home"),
    path("process-ai/", process_with_chatgpt, name="process_with_chatgpt"),
]

urlpatterns += static('/uploads/', document_root=os.path.join(settings.BASE_DIR, "uploads"))
urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
