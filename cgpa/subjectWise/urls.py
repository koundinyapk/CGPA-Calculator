from django.urls import path
from .views import *

urlpatterns=[
    path('',home),
    path("sub_wise",sub_wise)
]