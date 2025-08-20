from django.db import models
from django.contrib.auth.models import User


class Asset(models.Model):
	symbol = models.CharField(max_length=16, unique=True)
	name = models.CharField(max_length=64)
	is_active = models.BooleanField(default=True)

	def __str__(self):
		return self.symbol


class OptionContract(models.Model):
	CALL = 'CALL'
	PUT = 'PUT'
	DIRECTION_CHOICES = [(CALL, 'Call'), (PUT, 'Put')]

	asset = models.ForeignKey(Asset, on_delete=models.CASCADE, related_name='contracts')
	direction = models.CharField(max_length=4, choices=DIRECTION_CHOICES)
	expiry = models.DateTimeField()
	payout_rate = models.DecimalField(max_digits=5, decimal_places=2, help_text='Percent e.g. 80.00 for 80%')
	is_active = models.BooleanField(default=True)
	created_at = models.DateTimeField(auto_now_add=True)

	def __str__(self):
		return f"{self.asset.symbol} {self.direction} {self.expiry}"


class Trade(models.Model):
	PENDING = 'pending'
	WON = 'won'
	LOST = 'lost'
	CANCELLED = 'cancelled'
	STATUS_CHOICES = [
		(PENDING, 'Pending'),
		(WON, 'Won'),
		(LOST, 'Lost'),
		(CANCELLED, 'Cancelled'),
	]

	user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='trades')
	contract = models.ForeignKey(OptionContract, on_delete=models.PROTECT, related_name='trades')
	stake_amount = models.DecimalField(max_digits=12, decimal_places=2)
	status = models.CharField(max_length=16, choices=STATUS_CHOICES, default=PENDING)
	opened_at = models.DateTimeField(auto_now_add=True)
	settled_at = models.DateTimeField(null=True, blank=True)
	settlement_payout = models.DecimalField(max_digits=12, decimal_places=2, default=0)

	def __str__(self):
		return f"Trade({self.user.username}, {self.contract_id}, {self.status})"

# Create your models here.
