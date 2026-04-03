import 'package:flutter/material.dart';

class ProfileScreen extends StatelessWidget {
  const ProfileScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return ListView(
      padding: const EdgeInsets.all(16),
      children: const [
        _ProfileHeader(),
        SizedBox(height: 12),
        Card(
          child: Column(
            children: [
              ListTile(
                leading: Icon(Icons.person_outline),
                title: Text('Dados da conta'),
              ),
              Divider(height: 1),
              ListTile(
                leading: Icon(Icons.notifications_none),
                title: Text('Notificacoes'),
              ),
              Divider(height: 1),
              ListTile(
                leading: Icon(Icons.security_outlined),
                title: Text('Seguranca'),
              ),
            ],
          ),
        ),
      ],
    );
  }
}

class _ProfileHeader extends StatelessWidget {
  const _ProfileHeader();

  @override
  Widget build(BuildContext context) {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Row(
          children: const [
            CircleAvatar(
              radius: 28,
              child: Icon(Icons.person),
            ),
            SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    'Ismael',
                    style: TextStyle(fontWeight: FontWeight.w700, fontSize: 18),
                  ),
                  SizedBox(height: 4),
                  Text('Premium Plan'),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}
