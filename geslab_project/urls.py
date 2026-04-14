from django.contrib import admin
from django.urls import path, include
from muestras.views.auth import GesLabLoginView, GesLabLogoutView

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('muestras.urls'), name='muestras'),
    path('accounts/login/', GesLabLoginView.as_view(), name='login'),
    path('accounts/logout/', GesLabLogoutView.as_view(), name='logout'),
]
