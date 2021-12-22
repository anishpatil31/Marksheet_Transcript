from django.urls import path
from django.urls.resolvers import URLPattern
from . import views, proj2
from django.contrib import admin
from django.conf import settings
from django.conf.urls.static import static
urlpatterns=[
    path('', views.HomePage, name='HomePage'),
    path('run', proj2.solve, name='Proj2'),
    path('range', proj2.ranges, name='range')
]
urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)