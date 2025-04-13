import os

from django.conf import settings
from django.conf.urls.static import static
from django.urls import path, include

urlpatterns = [
    path("", include("processor.urls")),  # подключаем processor как корень
]

urlpatterns += static('/uploads/', document_root=os.path.join(settings.BASE_DIR, "uploads"))
urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
