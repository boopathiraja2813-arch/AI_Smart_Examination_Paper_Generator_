from django.urls import path
from . import views

urlpatterns = [
    path('', views.generate_paper_form, name='generate_paper_form'),
    path('chat/', views.chat_view, name='chat_view'),
    path('chat/reset/', views.chat_reset, name='chat_reset'),
    path('chat/finish/', views.chat_finish, name='chat_finish'),
]