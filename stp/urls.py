from django.conf.urls import url

from . import views

app_name = 'stp'
urlpatterns = [
	url(r'^$', views.getReport, name='getReport'),
]