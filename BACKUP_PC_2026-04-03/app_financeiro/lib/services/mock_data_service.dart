import 'package:app_financeiro/models/category_model.dart';
import 'package:app_financeiro/models/transaction_model.dart';
import 'package:flutter/material.dart';

class MockDataService {
  static List<CategoryModel> categories() {
    return const [
      CategoryModel(id: 'salary', name: 'Salario', icon: Icons.work_outline),
      CategoryModel(
        id: 'food',
        name: 'Alimentacao',
        icon: Icons.restaurant_outlined,
      ),
      CategoryModel(
        id: 'transport',
        name: 'Transporte',
        icon: Icons.directions_car_outlined,
      ),
      CategoryModel(
        id: 'health',
        name: 'Saude',
        icon: Icons.favorite_outline,
      ),
      CategoryModel(
        id: 'leisure',
        name: 'Lazer',
        icon: Icons.sports_esports_outlined,
      ),
    ];
  }

  static List<TransactionModel> transactions() {
    final now = DateTime.now();
    return [
      TransactionModel(
        id: 't1',
        title: 'Salario',
        amount: 6500,
        date: now.subtract(const Duration(days: 2)),
        type: TransactionType.income,
        category: 'salary',
      ),
      TransactionModel(
        id: 't2',
        title: 'Supermercado',
        amount: 240.80,
        date: now.subtract(const Duration(days: 1)),
        type: TransactionType.expense,
        category: 'food',
      ),
      TransactionModel(
        id: 't3',
        title: 'Combustivel',
        amount: 150.00,
        date: now.subtract(const Duration(days: 1)),
        type: TransactionType.expense,
        category: 'transport',
      ),
      TransactionModel(
        id: 't4',
        title: 'Consulta',
        amount: 110.00,
        date: now,
        type: TransactionType.expense,
        category: 'health',
      ),
    ];
  }
}
