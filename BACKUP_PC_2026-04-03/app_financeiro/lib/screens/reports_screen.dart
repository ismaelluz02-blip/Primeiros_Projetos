import 'package:app_financeiro/providers/transactions_provider.dart';
import 'package:app_financeiro/theme/app_theme.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

class ReportsScreen extends StatelessWidget {
  const ReportsScreen({super.key});

  String _money(double value) => 'R\$ ${value.toStringAsFixed(2)}';

  @override
  Widget build(BuildContext context) {
    final provider = context.watch<TransactionsProvider>();
    final income = provider.totalIncome;
    final expense = provider.totalExpense;
    final total = income + expense;
    final incomeRatio = total == 0 ? 0.0 : income / total;
    final expenseRatio = total == 0 ? 0.0 : expense / total;

    return ListView(
      padding: const EdgeInsets.all(16),
      children: [
        const Text(
          'Resumo',
          style: TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
        ),
        const SizedBox(height: 12),
        _summaryCard('Entradas', _money(income), AppTheme.darkGreen),
        const SizedBox(height: 10),
        _summaryCard('Saidas', _money(expense), AppTheme.expenseRed),
        const SizedBox(height: 10),
        _summaryCard('Saldo', _money(provider.balance), Colors.black87),
        const SizedBox(height: 16),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Text(
                  'Comparativo simples',
                  style: TextStyle(fontWeight: FontWeight.w600),
                ),
                const SizedBox(height: 12),
                _bar('Entradas', incomeRatio, AppTheme.darkGreen),
                const SizedBox(height: 10),
                _bar('Saidas', expenseRatio, AppTheme.expenseRed),
              ],
            ),
          ),
        ),
      ],
    );
  }

  Widget _summaryCard(String title, String value, Color color) {
    return Card(
      child: ListTile(
        title: Text(title),
        trailing: Text(
          value,
          style: TextStyle(fontWeight: FontWeight.w700, color: color),
        ),
      ),
    );
  }

  Widget _bar(String label, double ratio, Color color) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(label),
        const SizedBox(height: 6),
        ClipRRect(
          borderRadius: BorderRadius.circular(99),
          child: LinearProgressIndicator(
            value: ratio,
            minHeight: 10,
            backgroundColor: const Color(0xFFE7ECE7),
            valueColor: AlwaysStoppedAnimation<Color>(color),
          ),
        ),
      ],
    );
  }
}
