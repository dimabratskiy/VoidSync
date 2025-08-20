from django.urls import path
from .views import AssetListView, OptionContractListView, TradeCreateView, TradeSettleView


urlpatterns = [
	path('assets/', AssetListView.as_view(), name='assets'),
	path('contracts/', OptionContractListView.as_view(), name='contracts'),
	path('trades/', TradeCreateView.as_view(), name='trade_create'),
	path('trades/<int:trade_id>/settle/', TradeSettleView.as_view(), name='trade_settle'),
]
