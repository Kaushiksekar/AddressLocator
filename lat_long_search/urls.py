from django.urls import path

from . import views

app_name = 'lat_long_search'

urlpatterns = [
    path('', views.index, name='index'),
    path('get_excel', views.get_excel, name='get_excel'),
]