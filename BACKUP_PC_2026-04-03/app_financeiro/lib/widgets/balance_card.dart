import 'package:app_financeiro/theme/app_theme.dart';
import 'package:flutter/material.dart';

class BalanceCard extends StatelessWidget {
  const BalanceCard({
    super.key,
    required this.balance,
    required this.income,
    required this.expense,
    required this.isHidden,
    required this.onToggle,
  });

  final double balance;
  final double income;
  final double expense;
  final bool isHidden;
  final VoidCallback onToggle;

  String _money(double value) => 'R\$ ${value.toStringAsFixed(2)}';

  @override
  Widget build(BuildContext context) {
    final positive = balance >= 0;
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                const Text(
                  'Saldo total',
                  style: TextStyle(fontWeight: FontWeight.w600),
                ),
                const Spacer(),
                IconButton(
                  onPressed: onToggle,
                  icon: Icon(isHidden ? Icons.visibility_off : Icons.visibility),
                ),
              ],
            ),
            Text(
              isHidden ? 'R\$ ----' : _money(balance),
              style: const TextStyle(fontSize: 30, fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 6),
            Text(
              positive ? 'Voce esta no positivo' : 'Voce esta no negativo',
              style: TextStyle(
                color: positive ? AppTheme.darkGreen : AppTheme.expenseRed,
              ),
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                Expanded(
                  child: _statItem(
                    title: 'Entradas',
                    value: isHidden ? 'R\$ ----' : _money(income),
                    color: AppTheme.darkGreen,
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _statItem(
                    title: 'Saidas',
                    value: isHidden ? 'R\$ ----' : _money(expense),
                    color: AppTheme.expenseRed,
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _statItem({
    required String title,
    required String value,
    required Color color,
  }) {
    return Container(
      padding: const EdgeInsets.all(12),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: const Color(0xFFE3E9E3)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(title, style: const TextStyle(color: AppTheme.textSecondary)),
          const SizedBox(height: 4),
          Text(
            value,
            style: TextStyle(
              color: color,
              fontWeight: FontWeight.w700,
            ),
          ),
        ],
      ),
    );
  }
}
