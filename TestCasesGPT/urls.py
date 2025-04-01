from django.urls import include, path

urlpatterns = [
    path("processor/", include("processor.urls")),
]
