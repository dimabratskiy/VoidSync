from rest_framework import generics, permissions
from django.contrib.auth.models import User
from .serializers import RegistrationSerializer, ProfileSerializer, UserSerializer
from .models import Profile


class RegisterView(generics.CreateAPIView):
	queryset = User.objects.all()
	permission_classes = [permissions.AllowAny]
	serializer_class = RegistrationSerializer


class MeView(generics.RetrieveAPIView):
	serializer_class = UserSerializer

	def get_object(self):
		return self.request.user


class MyProfileView(generics.RetrieveUpdateAPIView):
	serializer_class = ProfileSerializer

	def get_object(self):
		return Profile.objects.get(user=self.request.user)

# Create your views here.
