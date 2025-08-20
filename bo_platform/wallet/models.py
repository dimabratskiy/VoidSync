from django.db import models
from django.contrib.auth.models import User
from django.core.validators import MinValueValidator


class Wallet(models.Model):
	user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='wallet')
	balance = models.DecimalField(max_digits=12, decimal_places=2, default=0)
	currency = models.CharField(max_length=10, default='USD')
	created_at = models.DateTimeField(auto_now_add=True)
	updated_at = models.DateTimeField(auto_now=True)

	def __str__(self):
		return f"Wallet({self.user.username}, {self.balance} {self.currency})"


class Transaction(models.Model):
	DEPOSIT = 'deposit'
	WITHDRAWAL = 'withdrawal'
	TRADE_SETTLEMENT = 'trade_settlement'
	TRANSACTION_TYPES = [
		(DEPOSIT, 'Deposit'),
		(WITHDRAWAL, 'Withdrawal'),
		(TRADE_SETTLEMENT, 'Trade Settlement'),
	]

	wallet = models.ForeignKey(Wallet, on_delete=models.CASCADE, related_name='transactions')
	amount = models.DecimalField(max_digits=12, decimal_places=2, validators=[MinValueValidator(0)])
	type = models.CharField(max_length=32, choices=TRANSACTION_TYPES)
	meta = models.JSONField(default=dict, blank=True)
	created_at = models.DateTimeField(auto_now_add=True)

	class Meta:
		ordering = ['-created_at']

	def __str__(self):
		return f"Transaction({self.type}, {self.amount})"

# Create your models here.
