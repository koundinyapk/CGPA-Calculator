from django.urls import path
from .views import *


urlpatterns=[
    path('',home),
    path("cgpa_calc",cgpa_cal)
]