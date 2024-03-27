import 'package:flutter/material.dart';

class BoxContainer extends StatelessWidget {
  final Widget child;

  const BoxContainer({
    super.key,
    required this.child,
  });
  @override
  Widget build(BuildContext context) {
    return Container(
      decoration: BoxDecoration(
        border: Border.all(
            color: Theme.of(context).colorScheme.primary.withOpacity(0.1)),
        borderRadius: BorderRadius.circular(8.0),
      ),
      child: child,
    );
  }
}
