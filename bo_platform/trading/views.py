from decimal import Decimal
from django.utils import timezone
from rest_framework import generics, permissions, status
from rest_framework.response import Response
from django.db import transaction as db_transaction
from .models import Asset, OptionContract, Trade
from .serializers import AssetSerializer, OptionContractSerializer, TradeSerializer
from wallet.models import Wallet, Transaction as WalletTransaction


class AssetListView(generics.ListAPIView):
	queryset = Asset.objects.filter(is_active=True)
	serializer_class = AssetSerializer
	permission_classes = [permissions.AllowAny]


class OptionContractListView(generics.ListAPIView):
	serializer_class = OptionContractSerializer

	def get_queryset(self):
		return OptionContract.objects.filter(is_active=True, expiry__gt=timezone.now())


class TradeCreateView(generics.CreateAPIView):
	serializer_class = TradeSerializer

	def create(self, request, *args, **kwargs):
		stake_amount = Decimal(str(request.data.get('stake_amount', '0')))
		contract_id = request.data.get('contract_id')
		if not contract_id or stake_amount <= 0:
			return Response({"detail": "contract_id and positive stake_amount required"}, status=status.HTTP_400_BAD_REQUEST)
		wallet, _ = Wallet.objects.get_or_create(user=request.user)
		if wallet.balance < stake_amount:
			return Response({"detail": "insufficient balance"}, status=status.HTTP_400_BAD_REQUEST)
		with db_transaction.atomic():
			contract = OptionContract.objects.select_for_update().get(pk=contract_id)
			trade = Trade.objects.create(user=request.user, contract=contract, stake_amount=stake_amount)
			wallet.balance = wallet.balance - stake_amount
			wallet.save(update_fields=["balance"])
			WalletTransaction.objects.create(wallet=wallet, amount=stake_amount, type=WalletTransaction.TRADE_SETTLEMENT, meta={"trade_id": trade.id, "action": "stake"})
		return Response(TradeSerializer(trade).data, status=status.HTTP_201_CREATED)


class TradeSettleView(generics.UpdateAPIView):
	serializer_class = TradeSerializer
	lookup_url_kwarg = 'trade_id'

	def update(self, request, *args, **kwargs):
		trade = Trade.objects.select_for_update().get(pk=kwargs['trade_id'], user=request.user)
		result = request.data.get('result')  # 'won' or 'lost' - in real life, decided by price feed
		if trade.status != Trade.PENDING:
			return Response({"detail": "trade already settled"}, status=status.HTTP_400_BAD_REQUEST)
		with db_transaction.atomic():
			if result == 'won':
				payout = (trade.stake_amount * trade.contract.payout_rate) / Decimal('100.0')
				trade.status = Trade.WON
				trade.settlement_payout = payout
				wallet, _ = Wallet.objects.get_or_create(user=request.user)
				wallet.balance = wallet.balance + trade.stake_amount + payout
				wallet.save(update_fields=["balance"])
				WalletTransaction.objects.create(wallet=wallet, amount=trade.stake_amount + payout, type=WalletTransaction.TRADE_SETTLEMENT, meta={"trade_id": trade.id, "action": "won"})
			else:
				trade.status = Trade.LOST
				trade.settlement_payout = Decimal('0')
			trade.settled_at = timezone.now()
			trade.save()
		return Response(TradeSerializer(trade).data)

# Create your views here.
