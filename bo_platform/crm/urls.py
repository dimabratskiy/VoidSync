from django.urls import path
from .views import TicketListCreateView, TicketDetailView, NoteCreateView


urlpatterns = [
	path('tickets/', TicketListCreateView.as_view(), name='tickets'),
	path('tickets/<int:pk>/', TicketDetailView.as_view(), name='ticket_detail'),
	path('notes/', NoteCreateView.as_view(), name='note_create'),
]
