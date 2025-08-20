from rest_framework import serializers
from .models import Asset, OptionContract, Trade


class AssetSerializer(serializers.ModelSerializer):
	class Meta:
		model = Asset
		fields = ["id", "symbol", "name", "is_active"]


class OptionContractSerializer(serializers.ModelSerializer):
	asset = AssetSerializer(read_only=True)
	asset_id = serializers.PrimaryKeyRelatedField(queryset=Asset.objects.all(), source='asset', write_only=True)

	class Meta:
		model = OptionContract
		fields = ["id", "asset", "asset_id", "direction", "expiry", "payout_rate", "is_active"]


class TradeSerializer(serializers.ModelSerializer):
	contract = OptionContractSerializer(read_only=True)
	contract_id = serializers.PrimaryKeyRelatedField(queryset=OptionContract.objects.all(), source='contract', write_only=True)

	class Meta:
		model = Trade
		fields = ["id", "contract", "contract_id", "stake_amount", "status", "opened_at", "settled_at", "settlement_payout"]
		read_only_fields = ["status", "opened_at", "settled_at", "settlement_payout"]
