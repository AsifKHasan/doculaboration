import 'package:flutter/material.dart';

class ElevatedCardWidget extends StatelessWidget {
  final Widget child;

  const ElevatedCardWidget({
    super.key,
    required this.child,
  });
  @override
  Widget build(BuildContext context) {
    final backgroundColor = Theme.of(context).colorScheme.surface;
    final primaryColor = Theme.of(context).colorScheme.primary;
    return Container(
      decoration: BoxDecoration(
        color: backgroundColor,
        boxShadow: [
          BoxShadow(
            color: primaryColor.withOpacity(0.1),
            blurRadius: 8,
          ),
        ],
        borderRadius: BorderRadius.circular(16.0),
      ),
      child: child,
    );
  }
}
