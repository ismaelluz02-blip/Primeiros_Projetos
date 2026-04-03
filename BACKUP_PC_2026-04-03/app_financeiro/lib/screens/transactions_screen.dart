import 'package:app_financeiro/providers/transactions_provider.dart';
import 'package:app_financeiro/widgets/transaction_tile.dart';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';

class TransactionsScreen extends StatelessWidget {
  const TransactionsScreen({super.key});

  @override
  Widget build(BuildContext context) {
    final provider = context.watch<TransactionsProvider>();
    final items = provider.transactions.toList()
      ..sort((a, b) => b.date.compareTo(a.date));

    return ListView.separated(
      padding: const EdgeInsets.all(16),
      itemCount: items.length,
      separatorBuilder: (_, _) => const SizedBox(height: 8),
      itemBuilder: (context, index) {
        final tx = items[index];
        return Card(
          child: TransactionTile(
            transaction: tx,
            category: provider.categoryById(tx.category),
          ),
        );
      },
    );
  }
}
