from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('test', views.test, name='test'),
    path('test2', views.test2, name='test2'),
    path('new_design', views.new_design, name='new_design'),
    path('download_template', views.download_template, name='download_template'),

]