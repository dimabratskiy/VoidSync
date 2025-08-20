from rest_framework import generics, permissions, status
from rest_framework.response import Response
from django.db import transaction as db_transaction
from .models import Wallet, Transaction
from .serializers import WalletSerializer, TransactionSerializer


class MyWalletView(generics.RetrieveAPIView):
	serializer_class = WalletSerializer

	def get_object(self):
		wallet, _ = Wallet.objects.get_or_create(user=self.request.user)
		return wallet


class TransactionListView(generics.ListAPIView):
	serializer_class = TransactionSerializer

	def get_queryset(self):
		wallet, _ = Wallet.objects.get_or_create(user=self.request.user)
		return wallet.transactions.all()


class DepositView(generics.CreateAPIView):
	serializer_class = TransactionSerializer
	permission_classes = [permissions.IsAuthenticated]

	def create(self, request, *args, **kwargs):
		amount = request.data.get('amount')
		if amount is None:
			return Response({"detail": "amount is required"}, status=status.HTTP_400_BAD_REQUEST)
		wallet, _ = Wallet.objects.get_or_create(user=request.user)
		with db_transaction.atomic():
			transaction = Transaction.objects.create(wallet=wallet, amount=amount, type=Transaction.DEPOSIT, meta={"method": "stub"})
			wallet.balance = wallet.balance + transaction.amount
			wallet.save(update_fields=["balance"])
		return Response(TransactionSerializer(transaction).data, status=status.HTTP_201_CREATED)

# Create your views here.
