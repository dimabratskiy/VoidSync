from django.contrib import admin
from .models import Wallet, Transaction


@admin.register(Wallet)
class WalletAdmin(admin.ModelAdmin):
	list_display = ("user", "balance", "currency", "created_at")
	search_fields = ("user__username", "user__email")
	list_filter = ("currency",)


@admin.register(Transaction)
class TransactionAdmin(admin.ModelAdmin):
	list_display = ("wallet", "type", "amount", "created_at")
	list_filter = ("type", "created_at")
	search_fields = ("wallet__user__username",)

# Register your models here.
