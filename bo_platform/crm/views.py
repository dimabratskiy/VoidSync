from rest_framework import generics
from .models import Ticket, Note
from .serializers import TicketSerializer, NoteSerializer


class TicketListCreateView(generics.ListCreateAPIView):
	serializer_class = TicketSerializer

	def get_queryset(self):
		return Ticket.objects.filter(user=self.request.user)

	def perform_create(self, serializer):
		serializer.save(user=self.request.user)


class TicketDetailView(generics.RetrieveUpdateAPIView):
	serializer_class = TicketSerializer

	def get_queryset(self):
		return Ticket.objects.filter(user=self.request.user)


class NoteCreateView(generics.CreateAPIView):
	serializer_class = NoteSerializer

	def perform_create(self, serializer):
		serializer.save(author=self.request.user)

# Create your views here.
