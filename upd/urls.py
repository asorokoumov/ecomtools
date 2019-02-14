from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='post_list'),
    path('download_template', views.download_template, name='download_template'),

]