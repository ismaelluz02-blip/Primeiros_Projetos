import 'package:app_financeiro/models/category_model.dart';
import 'package:app_financeiro/models/transaction_model.dart';
import 'package:app_financeiro/services/mock_data_service.dart';
import 'package:flutter/foundation.dart';

class TransactionsProvider extends ChangeNotifier {
  TransactionsProvider() {
    _categories = MockDataService.categories();
    _transactions = MockDataService.transactions();
  }

  late List<CategoryModel> _categories;
  late List<TransactionModel> _transactions;

  List<CategoryModel> get categories => List.unmodifiable(_categories);
  List<TransactionModel> get transactions => List.unmodifiable(_transactions);

  double get totalIncome => _transactions
      .where((t) => t.type == TransactionType.income)
      .fold(0.0, (sum, item) => sum + item.amount);

  double get totalExpense => _transactions
      .where((t) => t.type == TransactionType.expense)
      .fold(0.0, (sum, item) => sum + item.amount);

  double get balance => totalIncome - totalExpense;

  List<TransactionModel> get recentTransactions {
    final sorted = List<TransactionModel>.from(_transactions)
      ..sort((a, b) => b.date.compareTo(a.date));
    return sorted.take(5).toList();
  }

  void addTransaction(TransactionModel tx) {
    _transactions = [tx, ..._transactions];
    notifyListeners();
  }

  CategoryModel? categoryById(String id) {
    for (final category in _categories) {
      if (category.id == id) {
        return category;
      }
    }
    return null;
  }

  Map<String, double> expenseByCategory() {
    final data = <String, double>{};
    for (final tx in _transactions) {
      if (tx.type == TransactionType.expense) {
        data.update(tx.category, (v) => v + tx.amount, ifAbsent: () => tx.amount);
      }
    }
    return data;
  }
}
