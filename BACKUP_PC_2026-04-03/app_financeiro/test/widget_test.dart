import 'package:app_financeiro/main.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  testWidgets('Finance app shell renders', (WidgetTester tester) async {
    await tester.pumpWidget(const MyApp());

    expect(find.text('Ola, Ismael'), findsOneWidget);
    expect(find.text('Saldo total'), findsOneWidget);
    expect(find.text('Transacoes recentes'), findsOneWidget);
  });
}
