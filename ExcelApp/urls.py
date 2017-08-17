from django.conf.urls import url

from . import views

app_name = 'ExcelApp'
urlpatterns = [
    url(r'^$', views.index, name='index'),
    url(r'^details/$', views.details, name='details'),
    #url(r'^details/(?P<pid>[0-9]+)/$', views.getReport, name='getReport'),
    url(r'^getReport/$', views.getReport, name='getReport'),
    #url(r'^ExcelApp/$', views.details, name='details'),
]