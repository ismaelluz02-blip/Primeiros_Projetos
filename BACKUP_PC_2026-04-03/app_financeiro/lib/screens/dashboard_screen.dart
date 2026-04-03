import 'package:app_financeiro/providers/transactions_provider.dart';
import 'package:app_financeiro/widgets/balance_card.dart';
import 'package:app_financeiro/widgets/transaction_tile.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

class DashboardScreen extends StatefulWidget {
  const DashboardScreen({super.key});

  @override
  State<DashboardScreen> createState() => _DashboardScreenState();
}

class _DashboardScreenState extends State<DashboardScreen> {
  bool _hideBalance = false;

  @override
  Widget build(BuildContext context) {
    final provider = context.watch<TransactionsProvider>();
    final topCategory = _topExpenseCategory(provider);

    return ListView(
      padding: const EdgeInsets.all(16),
      children: [
        const Text(
          'Ola, Ismael',
          style: TextStyle(fontSize: 22, fontWeight: FontWeight.w700),
        ),
        const SizedBox(height: 4),
        const Text(
          'Resumo rapido da sua vida financeira',
          style: TextStyle(color: Color(0xFF7B867F)),
        ),
        const SizedBox(height: 16),
        BalanceCard(
          balance: provider.balance,
          income: provider.totalIncome,
          expense: provider.totalExpense,
          isHidden: _hideBalance,
          onToggle: () => setState(() => _hideBalance = !_hideBalance),
        ),
        const SizedBox(height: 12),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Text(
                  'Evolucao simples',
                  style: TextStyle(fontWeight: FontWeight.w600),
                ),
                const SizedBox(height: 12),
                Container(
                  height: 120,
                  decoration: BoxDecoration(
                    borderRadius: BorderRadius.circular(12),
                    color: const Color(0xFFF5F8F4),
                  ),
                  alignment: Alignment.center,
                  child: const Text('Grafico (placeholder)'),
                ),
              ],
            ),
          ),
        ),
        const SizedBox(height: 12),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Text(
              topCategory == null
                  ? 'Sem gastos suficientes para gerar insight.'
                  : 'Maior gasto recente: ${topCategory.$1} (R\$ ${topCategory.$2.toStringAsFixed(2)})',
            ),
          ),
        ),
        const SizedBox(height: 12),
        const Text(
          'Transacoes recentes',
          style: TextStyle(fontSize: 16, fontWeight: FontWeight.w700),
        ),
        const SizedBox(height: 8),
        ...provider.recentTransactions.map((tx) {
          return Card(
            child: TransactionTile(
              transaction: tx,
              category: provider.categoryById(tx.category),
            ),
          );
        }),
      ],
    );
  }

  (String, double)? _topExpenseCategory(TransactionsProvider provider) {
    final map = provider.expenseByCategory();
    if (map.isEmpty) return null;
    final sorted = map.entries.toList()..sort((a, b) => b.value.compareTo(a.value));
    final first = sorted.first;
    final category = provider.categoryById(first.key);
    return (category?.name ?? first.key, first.value);
  }
}
