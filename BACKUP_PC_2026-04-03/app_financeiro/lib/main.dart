import 'dart:math' as math;

import 'package:app_financeiro/screens/add_transaction_screen.dart';
import 'package:flutter/material.dart';

void main() {
  runApp(const MyApp());
}

class AppColors {
  static const Color primary = Color(0xFFA3D977);
  static const Color primarySoft = Color(0xFFDFF0CC);
  static const Color primaryDark = Color(0xFF2F5D35);
  static const Color background = Color(0xFFF7FAF5);
  static const Color card = Color(0xFFFFFFFF);
  static const Color textPrimary = Color(0xFF1D281E);
  static const Color textSecondary = Color(0xFF768276);
  static const Color border = Color(0xFFE3EADF);
  static const Color income = Color(0xFF4A8E53);
  static const Color expense = Color(0xFFE07E7E);
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Fintech Dashboard',
      theme: ThemeData(
        useMaterial3: true,
        scaffoldBackgroundColor: AppColors.background,
        colorScheme: ColorScheme.fromSeed(
          seedColor: AppColors.primary,
          brightness: Brightness.light,
        ),
      ),
      home: const FinanceDashboardPage(),
    );
  }
}

class FinanceDashboardPage extends StatefulWidget {
  const FinanceDashboardPage({super.key});

  @override
  State<FinanceDashboardPage> createState() => _FinanceDashboardPageState();
}

class _FinanceDashboardPageState extends State<FinanceDashboardPage> {
  bool _showBalance = true;
  List<Map<String, dynamic>> transactions = [];

  void _toggleBalance() {
    setState(() {
      _showBalance = !_showBalance;
    });
  }

  Future<void> _openAddTransaction() async {
    final result = await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => const AddTransactionScreen(),
      ),
    );

    if (result != null) {
      final tx = Map<String, dynamic>.from(result as Map);
      tx['createdAt'] = DateTime.now();
      setState(() {
        transactions.insert(0, tx);
      });
    }
  }

  List<double> _weekExpenseTotals() {
    final now = DateTime.now();
    final today = DateTime(now.year, now.month, now.day);
    final startOfWeek = today.subtract(Duration(days: today.weekday - 1));
    final endOfWeek = startOfWeek.add(const Duration(days: 7));
    final totals = List<double>.filled(7, 0);

    for (final tx in transactions) {
      if (tx['type'] != 'Expense') {
        continue;
      }

      final rawDate = tx['createdAt'];
      DateTime txDate;
      if (rawDate is DateTime) {
        txDate = DateTime(rawDate.year, rawDate.month, rawDate.day);
      } else {
        txDate = today;
      }

      if (txDate.isBefore(startOfWeek) || !txDate.isBefore(endOfWeek)) {
        continue;
      }

      final amount = _parseAmount(tx['amount']);
      final index = txDate.weekday - 1;
      totals[index] += amount;
    }

    return totals;
  }

  double _parseAmount(dynamic rawValue) {
    if (rawValue == null) {
      return 0;
    }

    var value = rawValue.toString().trim();
    value = value.replaceAll('R\$', '').replaceAll(' ', '');

    if (value.contains(',') && value.contains('.')) {
      value = value.replaceAll('.', '').replaceAll(',', '.');
    } else if (value.contains(',')) {
      value = value.replaceAll(',', '.');
    }

    return double.tryParse(value) ?? 0;
  }

  @override
  Widget build(BuildContext context) {
    final weekTotals = _weekExpenseTotals();

    return Scaffold(
      body: Stack(
        children: [
          Positioned(
            top: -180,
            left: -140,
            child: Container(
              width: 360,
              height: 360,
              decoration: const BoxDecoration(
                shape: BoxShape.circle,
                gradient: RadialGradient(
                  colors: [Color(0x66CDE8AE), Color(0x00CDE8AE)],
                ),
              ),
            ),
          ),
          SafeArea(
            child: ListView(
              padding: const EdgeInsets.fromLTRB(16, 10, 16, 100),
              children: [
                _buildHeader(),
                const SizedBox(height: 16),
                _BalanceHeroCard(
                  showBalance: _showBalance,
                  onToggle: _toggleBalance,
                ),
                const SizedBox(height: 20),
                const _SectionTitle('Resumo rapido'),
                const SizedBox(height: 12),
                _buildSummaryCards(),
                const SizedBox(height: 20),
                const _SectionTitle('Indicadores da semana'),
                const SizedBox(height: 12),
                _WeeklyExpenseCard(totals: weekTotals),
                const SizedBox(height: 20),
                const InsightCard(
                  message: 'Parabens! Voce gastou menos do que recebeu este mes.',
                ),
                const SizedBox(height: 22),
                const _SectionTitle('Transacoes recentes'),
                const SizedBox(height: 12),
                if (transactions.isEmpty)
                  const _EmptyTransactionsCard()
                else
                  Column(
                    children: transactions.map((t) {
                      final isExpense = t['type'] == 'Expense';
                      final amount = t['amount']?.toString() ?? '0';
                      final category = t['category']?.toString() ?? 'Sem categoria';
                      final type = t['type']?.toString() ?? 'Type';

                      return Padding(
                        padding: const EdgeInsets.only(bottom: 12),
                        child: TransactionTile(
                          icon: isExpense ? Icons.arrow_upward_rounded : Icons.arrow_downward_rounded,
                          title: category,
                          subtitle: type,
                          amount: '${isExpense ? '-' : '+'} R\$ $amount',
                          amountColor: isExpense ? AppColors.expense : AppColors.income,
                        ),
                      );
                    }).toList(),
                  ),
              ],
            ),
          ),
        ],
      ),
      floatingActionButton: Container(
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(18),
          boxShadow: const [
            BoxShadow(
              color: Color(0x29000000),
              blurRadius: 20,
              offset: Offset(0, 10),
            ),
          ],
        ),
        child: FloatingActionButton(
          onPressed: _openAddTransaction,
          backgroundColor: AppColors.primary,
          foregroundColor: AppColors.primaryDark,
          elevation: 0,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(18),
          ),
          child: const Icon(Icons.add, size: 30),
        ),
      ),
    );
  }

  Widget _buildHeader() {
    return Row(
      children: [
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: const [
              Text(
                'Ola, Ismael',
                style: TextStyle(
                  color: AppColors.textPrimary,
                  fontWeight: FontWeight.w800,
                  fontSize: 27,
                  letterSpacing: -0.5,
                  height: 1.05,
                ),
              ),
              SizedBox(height: 5),
              Text(
                'Sua visao financeira de hoje',
                style: TextStyle(
                  color: AppColors.textSecondary,
                  fontWeight: FontWeight.w500,
                ),
              ),
            ],
          ),
        ),
        _HeaderAction(
          onTap: _toggleBalance,
          icon: _showBalance ? Icons.visibility_off_outlined : Icons.visibility_outlined,
        ),
        const SizedBox(width: 8),
        const CircleAvatar(
          radius: 20,
          backgroundColor: Color(0xFFE9F4DF),
          child: Icon(Icons.person_outline, color: AppColors.primaryDark),
        ),
      ],
    );
  }

  Widget _buildSummaryCards() {
    return SizedBox(
      height: 130,
      child: ListView(
        scrollDirection: Axis.horizontal,
        children: const [
          SummaryCard(
            title: 'Income',
            value: 'R\$ 14.300,00',
            icon: Icons.south_west_rounded,
            valueColor: AppColors.income,
            gradient: [Color(0xFFF5FCED), Color(0xFFEAF7DC)],
          ),
          SizedBox(width: 12),
          SummaryCard(
            title: 'Expenses',
            value: 'R\$ 1.849,10',
            icon: Icons.north_east_rounded,
            valueColor: AppColors.expense,
            gradient: [Color(0xFFFFF5F5), Color(0xFFFFEBEB)],
          ),
          SizedBox(width: 12),
          SummaryCard(
            title: 'Balance',
            value: 'R\$ 12.450,90',
            icon: Icons.account_balance_wallet_rounded,
            valueColor: AppColors.income,
            gradient: [Color(0xFFF4FBEC), Color(0xFFE7F5D6)],
          ),
        ],
      ),
    );
  }
}

class _HeaderAction extends StatelessWidget {
  const _HeaderAction({
    required this.onTap,
    required this.icon,
  });

  final VoidCallback onTap;
  final IconData icon;

  @override
  Widget build(BuildContext context) {
    return InkWell(
      onTap: onTap,
      borderRadius: BorderRadius.circular(14),
      child: Ink(
        width: 42,
        height: 42,
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(14),
          border: Border.all(color: AppColors.border),
          boxShadow: const [
            BoxShadow(
              color: Color(0x0E000000),
              blurRadius: 10,
              offset: Offset(0, 5),
            ),
          ],
        ),
        child: Icon(icon, size: 21, color: AppColors.textSecondary),
      ),
    );
  }
}

class _BalanceHeroCard extends StatelessWidget {
  const _BalanceHeroCard({
    required this.showBalance,
    required this.onToggle,
  });

  final bool showBalance;
  final VoidCallback onToggle;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.fromLTRB(20, 18, 20, 18),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(30),
        border: Border.all(color: const Color(0xDCE2EADB)),
        gradient: const LinearGradient(
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
          colors: [Color(0xFFFFFFFF), Color(0xFFF0F8E8)],
        ),
        boxShadow: const [
          BoxShadow(
            color: Color(0x18000000),
            blurRadius: 28,
            offset: Offset(0, 14),
          ),
        ],
      ),
      child: Stack(
        children: [
          Positioned(
            right: -42,
            top: -32,
            child: Container(
              width: 140,
              height: 140,
              decoration: const BoxDecoration(
                shape: BoxShape.circle,
                gradient: RadialGradient(
                  colors: [Color(0x55B8E08B), Color(0x00B8E08B)],
                ),
              ),
            ),
          ),
          Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                children: [
                  const Text(
                    'Saldo total',
                    style: TextStyle(
                      color: AppColors.textSecondary,
                      fontSize: 14,
                      fontWeight: FontWeight.w600,
                    ),
                  ),
                  const Spacer(),
                  InkWell(
                    onTap: onToggle,
                    borderRadius: BorderRadius.circular(999),
                    child: Ink(
                      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                      decoration: BoxDecoration(
                        color: Colors.white.withValues(alpha: 0.88),
                        borderRadius: BorderRadius.circular(999),
                        border: Border.all(color: AppColors.border),
                      ),
                      child: Row(
                        children: [
                          Icon(
                            showBalance
                                ? Icons.visibility_off_outlined
                                : Icons.visibility_outlined,
                            size: 16,
                            color: AppColors.primaryDark,
                          ),
                          const SizedBox(width: 6),
                          Text(
                            showBalance ? 'Ocultar' : 'Mostrar',
                            style: const TextStyle(
                              color: AppColors.primaryDark,
                              fontWeight: FontWeight.w600,
                              fontSize: 12,
                            ),
                          ),
                        ],
                      ),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 14),
              Text(
                showBalance ? 'R\$ 12.450,90' : 'R\$ -------',
                style: const TextStyle(
                  fontSize: 42,
                  height: 1.02,
                  letterSpacing: -1.0,
                  fontWeight: FontWeight.w800,
                  color: AppColors.textPrimary,
                ),
              ),
              const SizedBox(height: 10),
              const Text(
                'Atualizado agora',
                style: TextStyle(
                  fontSize: 11,
                  color: AppColors.textSecondary,
                  fontWeight: FontWeight.w500,
                  letterSpacing: 0.2,
                ),
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _SectionTitle extends StatelessWidget {
  const _SectionTitle(this.title);

  final String title;

  @override
  Widget build(BuildContext context) {
    return Text(
      title,
      style: const TextStyle(
        fontSize: 19,
        color: AppColors.textPrimary,
        fontWeight: FontWeight.w700,
        letterSpacing: -0.25,
      ),
    );
  }
}

class SummaryCard extends StatelessWidget {
  const SummaryCard({
    super.key,
    required this.title,
    required this.value,
    required this.icon,
    required this.valueColor,
    required this.gradient,
  });

  final String title;
  final String value;
  final IconData icon;
  final Color valueColor;
  final List<Color> gradient;

  @override
  Widget build(BuildContext context) {
    return Container(
      width: 174,
      padding: const EdgeInsets.all(14),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: AppColors.border),
        gradient: LinearGradient(
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
          colors: gradient,
        ),
        boxShadow: const [
          BoxShadow(
            color: Color(0x0E000000),
            blurRadius: 12,
            offset: Offset(0, 6),
          ),
        ],
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(
            width: 32,
            height: 32,
            decoration: BoxDecoration(
              color: Colors.white.withValues(alpha: 0.86),
              borderRadius: BorderRadius.circular(11),
            ),
            child: Icon(icon, size: 18, color: valueColor),
          ),
          const Spacer(),
          Text(
            title,
            style: const TextStyle(
              fontSize: 12,
              color: AppColors.textSecondary,
              fontWeight: FontWeight.w600,
            ),
          ),
          const SizedBox(height: 5),
          Text(
            value,
            maxLines: 1,
            overflow: TextOverflow.ellipsis,
            style: TextStyle(
              fontSize: 14,
              fontWeight: FontWeight.w800,
              color: valueColor,
            ),
          ),
        ],
      ),
    );
  }
}

class InsightCard extends StatelessWidget {
  const InsightCard({super.key, required this.message});

  final String message;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: const Color(0xFFEFF8E6),
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: AppColors.border),
      ),
      child: Row(
        children: [
          Container(
            width: 36,
            height: 36,
            decoration: const BoxDecoration(
              shape: BoxShape.circle,
              color: Color(0xFFDDEECE),
            ),
            child: const Icon(Icons.auto_graph_rounded, color: AppColors.primaryDark),
          ),
          const SizedBox(width: 12),
          Expanded(
            child: Text(
              message,
              style: const TextStyle(
                fontSize: 13.5,
                height: 1.36,
                color: Color(0xFF355E35),
                fontWeight: FontWeight.w600,
              ),
            ),
          ),
          const SizedBox(width: 6),
          const Icon(Icons.chevron_right_rounded, color: AppColors.primaryDark),
        ],
      ),
    );
  }
}

class _WeeklyExpenseCard extends StatelessWidget {
  const _WeeklyExpenseCard({required this.totals});

  final List<double> totals;

  @override
  Widget build(BuildContext context) {
    const weekDays = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab', 'Dom'];
    final maxValue = totals.fold<double>(0, (acc, item) => math.max(acc, item));
    final totalWeek = totals.fold<double>(0, (acc, item) => acc + item);
    final topIndex = totals.indexOf(maxValue);

    final subtitle = totalWeek == 0
        ? 'Sem gastos registrados nesta semana'
        : 'Maior gasto: ${weekDays[topIndex]} (R\$ ${maxValue.toStringAsFixed(0)})';

    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        color: AppColors.card,
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: AppColors.border),
        boxShadow: const [
          BoxShadow(
            color: Color(0x0C000000),
            blurRadius: 14,
            offset: Offset(0, 7),
          ),
        ],
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            children: [
              Container(
                width: 34,
                height: 34,
                decoration: BoxDecoration(
                  color: AppColors.primarySoft,
                  borderRadius: BorderRadius.circular(11),
                ),
                child: const Icon(Icons.bar_chart_rounded, color: AppColors.primaryDark),
              ),
              const SizedBox(width: 10),
              const Text(
                'Gastos por dia',
                style: TextStyle(
                  color: AppColors.textPrimary,
                  fontWeight: FontWeight.w700,
                  fontSize: 15,
                ),
              ),
              const Spacer(),
              Text(
                'R\$ ${totalWeek.toStringAsFixed(0)}',
                style: const TextStyle(
                  color: AppColors.expense,
                  fontWeight: FontWeight.w700,
                ),
              ),
            ],
          ),
          const SizedBox(height: 8),
          Text(
            subtitle,
            style: const TextStyle(
              color: AppColors.textSecondary,
              fontSize: 12.5,
              fontWeight: FontWeight.w500,
            ),
          ),
          const SizedBox(height: 14),
          SizedBox(
            height: 146,
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.end,
              children: List.generate(7, (index) {
                return Expanded(
                  child: _WeekBar(
                    day: weekDays[index],
                    value: totals[index],
                    maxValue: maxValue,
                  ),
                );
              }),
            ),
          ),
        ],
      ),
    );
  }
}

class _WeekBar extends StatelessWidget {
  const _WeekBar({
    required this.day,
    required this.value,
    required this.maxValue,
  });

  final String day;
  final double value;
  final double maxValue;

  @override
  Widget build(BuildContext context) {
    final hasValue = value > 0;
    final height = maxValue <= 0 ? 8.0 : math.max(8.0, (value / maxValue) * 92);

    return Column(
      mainAxisAlignment: MainAxisAlignment.end,
      children: [
        Text(
          hasValue ? 'R\$ ${value.toStringAsFixed(0)}' : '',
          style: const TextStyle(
            fontSize: 10.5,
            color: AppColors.textSecondary,
            fontWeight: FontWeight.w600,
          ),
        ),
        const SizedBox(height: 6),
        Container(
          width: 22,
          height: 96,
          alignment: Alignment.bottomCenter,
          child: AnimatedContainer(
            duration: const Duration(milliseconds: 280),
            curve: Curves.easeOutCubic,
            height: height,
            decoration: BoxDecoration(
              borderRadius: BorderRadius.circular(8),
              gradient: LinearGradient(
                begin: Alignment.topCenter,
                end: Alignment.bottomCenter,
                colors: hasValue
                    ? const [Color(0xFFB9E38F), Color(0xFF7DBE5E)]
                    : const [Color(0xFFE8EEE3), Color(0xFFDDE5D8)],
              ),
            ),
          ),
        ),
        const SizedBox(height: 8),
        Text(
          day,
          style: const TextStyle(
            fontSize: 11.5,
            color: AppColors.textSecondary,
            fontWeight: FontWeight.w600,
          ),
        ),
      ],
    );
  }
}

class _EmptyTransactionsCard extends StatelessWidget {
  const _EmptyTransactionsCard();

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: AppColors.border),
        boxShadow: const [
          BoxShadow(
            color: Color(0x0B000000),
            blurRadius: 12,
            offset: Offset(0, 6),
          ),
        ],
      ),
      child: Column(
        children: [
          Container(
            width: 46,
            height: 46,
            decoration: BoxDecoration(
              color: const Color(0xFFF2F8ED),
              borderRadius: BorderRadius.circular(14),
            ),
            child: const Icon(Icons.receipt_long_outlined, color: AppColors.primaryDark),
          ),
          const SizedBox(height: 12),
          const Text(
            'Nenhuma transacao adicionada ainda',
            style: TextStyle(
              color: AppColors.textPrimary,
              fontWeight: FontWeight.w700,
            ),
          ),
          const SizedBox(height: 4),
          const Text(
            'Toque no botao + para adicionar sua primeira transacao.',
            textAlign: TextAlign.center,
            style: TextStyle(
              color: AppColors.textSecondary,
              fontSize: 12.5,
              height: 1.35,
            ),
          ),
        ],
      ),
    );
  }
}

class TransactionTile extends StatelessWidget {
  const TransactionTile({
    super.key,
    required this.icon,
    required this.title,
    required this.subtitle,
    required this.amount,
    required this.amountColor,
  });

  final IconData icon;
  final String title;
  final String subtitle;
  final String amount;
  final Color amountColor;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 14),
      decoration: BoxDecoration(
        color: AppColors.card,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: AppColors.border),
        boxShadow: const [
          BoxShadow(
            color: Color(0x0C000000),
            blurRadius: 14,
            offset: Offset(0, 7),
          ),
        ],
      ),
      child: Row(
        children: [
          Container(
            width: 39,
            height: 39,
            decoration: BoxDecoration(
              color: const Color(0xFFF1F7EC),
              borderRadius: BorderRadius.circular(12),
            ),
            child: Icon(icon, size: 20, color: AppColors.primaryDark),
          ),
          const SizedBox(width: 12),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  title,
                  style: const TextStyle(
                    fontSize: 14.5,
                    color: AppColors.textPrimary,
                    fontWeight: FontWeight.w700,
                  ),
                ),
                const SizedBox(height: 4),
                Text(
                  subtitle,
                  style: const TextStyle(
                    fontSize: 12,
                    color: AppColors.textSecondary,
                    fontWeight: FontWeight.w500,
                  ),
                ),
              ],
            ),
          ),
          const SizedBox(width: 10),
          Text(
            amount,
            style: TextStyle(
              fontSize: 14.5,
              color: amountColor,
              fontWeight: FontWeight.w800,
              letterSpacing: -0.1,
            ),
          ),
        ],
      ),
    );
  }
}
