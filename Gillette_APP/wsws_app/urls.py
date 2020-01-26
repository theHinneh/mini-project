from django.urls import path

from wsws_app import views

urlpatterns = [
    path('', views.FileView.as_view(), name='FileUpload'),
]
