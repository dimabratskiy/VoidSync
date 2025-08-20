from django.db import models
from django.contrib.auth.models import User


class Ticket(models.Model):
	OPEN = 'open'
	IN_PROGRESS = 'in_progress'
	RESOLVED = 'resolved'
	CLOSED = 'closed'
	STATUS_CHOICES = [
		(OPEN, 'Open'),
		(IN_PROGRESS, 'In Progress'),
		(RESOLVED, 'Resolved'),
		(CLOSED, 'Closed'),
	]

	user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='tickets')
	title = models.CharField(max_length=200)
	description = models.TextField()
	status = models.CharField(max_length=20, choices=STATUS_CHOICES, default=OPEN)
	created_at = models.DateTimeField(auto_now_add=True)
	updated_at = models.DateTimeField(auto_now=True)

	def __str__(self):
		return f"Ticket({self.title}, {self.status})"


class Note(models.Model):
	ticket = models.ForeignKey(Ticket, on_delete=models.CASCADE, related_name='notes')
	author = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, related_name='crm_notes')
	content = models.TextField()
	created_at = models.DateTimeField(auto_now_add=True)

	def __str__(self):
		return f"Note({self.ticket_id})"

# Create your models here.
