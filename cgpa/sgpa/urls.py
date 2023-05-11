from django.urls import path
from .views import *

urlpatterns=[
    path('',home),
    path('sgpa_calc',sgpa)
]