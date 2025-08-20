from django.urls import path
from .views import RegisterView, MeView, MyProfileView


urlpatterns = [
	path('register/', RegisterView.as_view(), name='register'),
	path('me/', MeView.as_view(), name='me'),
	path('profile/', MyProfileView.as_view(), name='profile'),
]
