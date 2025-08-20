from django.contrib import admin
from .models import Ticket, Note


@admin.register(Ticket)
class TicketAdmin(admin.ModelAdmin):
	list_display = ("id", "user", "title", "status", "created_at")
	list_filter = ("status", "created_at")
	search_fields = ("title", "description", "user__username", "user__email")


@admin.register(Note)
class NoteAdmin(admin.ModelAdmin):
	list_display = ("id", "ticket", "author", "created_at")
	search_fields = ("ticket__title", "author__username")

# Register your models here.
