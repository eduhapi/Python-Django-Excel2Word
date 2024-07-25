"""
URL configuration for excel_to_word_app project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.urls import path
from word_processor import views
from django.conf import settings
from django.conf.urls.static import static



urlpatterns = [
    path('upload/', views.upload_excel, name='upload_excel'),
    path('', views.upload_excel, name='upload_excel'),  # Map empty path to upload_excel view
    #path('merge_excel_success', views.merge_excel_files, name='merge_excel_success'),
    path('download/<str:filename>/', views.download_document, name='download_document'),
]
if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
