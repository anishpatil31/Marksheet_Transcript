from django.contrib import admin
from django.conf import settings
from django.conf.urls.static import static
from django.urls import path
from . import views, proj1, generate
urlpatterns = [
    path('', views.HomePage, name='Home'),
    path('submit', proj1.projf, name='Proj1'),
    path('GCM', generate.gcm, name='GCM'),
    path('send', views.send, name='send')
]

urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)


