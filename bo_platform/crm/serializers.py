from rest_framework import serializers
from .models import Ticket, Note


class TicketSerializer(serializers.ModelSerializer):
	class Meta:
		model = Ticket
		fields = ["id", "user", "title", "description", "status", "created_at", "updated_at"]
		read_only_fields = ["user", "created_at", "updated_at"]


class NoteSerializer(serializers.ModelSerializer):
	class Meta:
		model = Note
		fields = ["id", "ticket", "author", "content", "created_at"]
		read_only_fields = ["author", "created_at"]
