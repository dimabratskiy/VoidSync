from django.contrib import admin
from .models import Asset, OptionContract, Trade


@admin.register(Asset)
class AssetAdmin(admin.ModelAdmin):
	list_display = ("symbol", "name", "is_active")
	search_fields = ("symbol", "name")
	list_filter = ("is_active",)


@admin.register(OptionContract)
class OptionContractAdmin(admin.ModelAdmin):
	list_display = ("asset", "direction", "expiry", "payout_rate", "is_active")
	list_filter = ("direction", "is_active", "expiry")
	search_fields = ("asset__symbol",)


@admin.register(Trade)
class TradeAdmin(admin.ModelAdmin):
	list_display = ("user", "contract", "stake_amount", "status", "opened_at", "settled_at")
	list_filter = ("status", "opened_at", "settled_at")
	search_fields = ("user__username",)

# Register your models here.
