import 'package:flutter/material.dart';

class AppTheme {
  static const Color pearGreen = Color(0xFF9FD66A);
  static const Color darkGreen = Color(0xFF2E6B3B);
  static const Color bg = Color(0xFFFFFFFF);
  static const Color bgAlt = Color(0xFFF7FAF5);
  static const Color textPrimary = Color(0xFF2F3A33);
  static const Color textSecondary = Color(0xFF7B867F);
  static const Color expenseRed = Color(0xFFE57373);

  static ThemeData get lightTheme {
    final colorScheme = ColorScheme.fromSeed(
      seedColor: pearGreen,
      brightness: Brightness.light,
    ).copyWith(
      primary: pearGreen,
      secondary: darkGreen,
      error: expenseRed,
      surface: bg,
    );

    return ThemeData(
      useMaterial3: true,
      colorScheme: colorScheme,
      scaffoldBackgroundColor: bgAlt,
      appBarTheme: const AppBarTheme(
        centerTitle: false,
        backgroundColor: bg,
        foregroundColor: textPrimary,
        elevation: 0,
      ),
      cardTheme: CardThemeData(
        color: bg,
        elevation: 2,
        shadowColor: Colors.black12,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(16),
        ),
      ),
      inputDecorationTheme: InputDecorationTheme(
        filled: true,
        fillColor: bg,
        border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(14),
          borderSide: BorderSide.none,
        ),
        focusedBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(14),
          borderSide: const BorderSide(color: pearGreen),
        ),
      ),
      bottomNavigationBarTheme: const BottomNavigationBarThemeData(
        selectedItemColor: darkGreen,
        unselectedItemColor: textSecondary,
        type: BottomNavigationBarType.fixed,
      ),
    );
  }
}
