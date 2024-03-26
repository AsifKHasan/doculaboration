import 'package:flutter/material.dart';
import 'package:gap/gap.dart';

class LandingScreen extends StatefulWidget {
  final List<DrawerItem> items;
  final int initialSelectedItemIndex;
  const LandingScreen({
    super.key,
    required this.items,
    this.initialSelectedItemIndex = 0,
  });

  @override
  State<LandingScreen> createState() => _LandingScreenState();
}

class _LandingScreenState extends State<LandingScreen> {
  int _selectedItemIndex = 0;
  String title = "";
  Color backgroundColor = Colors.white;
  Color primaryColor = Colors.white;
  @override
  void initState() {
    super.initState();
    setState(() {
      _selectedItemIndex = widget.initialSelectedItemIndex;
      title = widget.items[_selectedItemIndex].title;
    });
  }

  @override
  void didChangeDependencies() {
    super.didChangeDependencies();
    setState(() {
      backgroundColor = Theme.of(context).colorScheme.background;
      primaryColor = Theme.of(context).colorScheme.primary;
    });
  }

  @override
  Widget build(BuildContext context) {
    return _largeView();
  }

  Widget _largeView() {
    return Scaffold(
      appBar: AppBar(title: Text(title)),
      body: SafeArea(
        child: Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            SizedBox(
              width: 200,
              child: LeftSectionWidget(
                items: widget.items,
                onItemSelected: (index) {
                  setState(() {
                    _selectedItemIndex = index;
                    title = widget.items[index].title;
                  });
                },
                selectedIndex: _selectedItemIndex,
                defaultColor: backgroundColor,
                selectedColor: primaryColor,
              ),
            ),
            const Gap(16),
            SizedBox(
              width: 800,
              child: widget.items[_selectedItemIndex].widget,
            ),
            const Gap(16),
            Flexible(
              // TODO: use it later for Jobs and Job Log section
              child: Container(width: 200),
            ),
          ],
        ),
      ),
    );
  }
}

class LeftSectionWidget extends StatelessWidget {
  final List<DrawerItem> items;
  final Function(int) onItemSelected;
  final int selectedIndex;
  final Color defaultColor;
  final Color selectedColor;

  const LeftSectionWidget({
    super.key,
    required this.items,
    required this.onItemSelected,
    required this.selectedColor,
    required this.defaultColor,
    required this.selectedIndex,
  });

  @override
  Widget build(BuildContext context) {
    return ListView.separated(
      itemBuilder: (cxt, index) {
        final item = items[index];
        return ListTile(
          selected: index == selectedIndex,
          selectedColor: selectedColor,
          tileColor: defaultColor,
          title: Text(item.title),
          leading: item.icon,
          onTap: () => onItemSelected.call(index),
        );
      },
      separatorBuilder: (cxt, index) {
        return const Gap(16.0);
      },
      itemCount: items.length,
    );
  }
}

class DrawerItem {
  final String title;
  final Widget widget;
  final Widget icon;

  DrawerItem({
    required this.title,
    required this.widget,
    required this.icon,
  });
}

class ElevatedCardWidget extends StatelessWidget {
  final Widget child;

  const ElevatedCardWidget({
    super.key,
    required this.child,
  });
  @override
  Widget build(BuildContext context) {
    final backgroundColor = Theme.of(context).colorScheme.background;
    final primaryColor = Theme.of(context).colorScheme.primary;
    return Container(
      decoration: BoxDecoration(
        color: backgroundColor,
        boxShadow: [
          BoxShadow(
            color: primaryColor.withOpacity(0.1),
            blurRadius: 5,
          ),
        ],
        borderRadius: BorderRadius.circular(16.0),
      ),
      child: child,
    );
  }
}
