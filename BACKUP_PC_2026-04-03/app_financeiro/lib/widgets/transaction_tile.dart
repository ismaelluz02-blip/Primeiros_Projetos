import 'package:app_financeiro/models/category_model.dart';
import 'package:app_financeiro/models/transaction_model.dart';
import 'package:app_financeiro/theme/app_theme.dart';
import 'package:flutter/material.dart';

class TransactionTile extends StatelessWidget {
  const TransactionTile({
    super.key,
    required this.transaction,
    required this.category,
  });

  final TransactionModel transaction;
  final CategoryModel? category;

  String _money(double value) => 'R\$ ${value.toStringAsFixed(2)}';

  @override
  Widget build(BuildContext context) {
    final isIncome = transaction.type == TransactionType.income;
    return ListTile(
      contentPadding: const EdgeInsets.symmetric(horizontal: 4, vertical: 4),
      leading: CircleAvatar(
        backgroundColor: const Color(0xFFF1F6EF),
        child: Icon(
          category?.icon ?? Icons.category_outlined,
          color: AppTheme.darkGreen,
        ),
      ),
      title: Text(transaction.title),
      subtitle: Text(
        '${category?.name ?? 'Sem categoria'} - ${transaction.date.day}/${transaction.date.month}',
      ),
      trailing: Text(
        '${isIncome ? '+' : '-'} ${_money(transaction.amount)}',
        style: TextStyle(
          color: isIncome ? AppTheme.darkGreen : AppTheme.expenseRed,
          fontWeight: FontWeight.w700,
        ),
      ),
    );
  }
}
