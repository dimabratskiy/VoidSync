from django.urls import path
from .views import MyWalletView, TransactionListView, DepositView


urlpatterns = [
	path('me/', MyWalletView.as_view(), name='my_wallet'),
	path('transactions/', TransactionListView.as_view(), name='wallet_transactions'),
	path('deposit/', DepositView.as_view(), name='wallet_deposit'),
]
