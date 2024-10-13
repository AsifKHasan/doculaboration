import 'package:flutter/material.dart';
import 'package:flutter_bloc/flutter_bloc.dart';
import 'package:gap/gap.dart';
import 'package:ui/feature/gseet_input/util/job_state_stage.dart';

import '../gseet_input/presentation/log_widget.dart';
import '../shared/presentation/bloc/job_queue_bloc.dart';
import '../shared/presentation/bloc/selected_item_bloc.dart';

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
      backgroundColor = Theme.of(context).colorScheme.surface;
      primaryColor = Theme.of(context).colorScheme.primary;
    });
  }

  @override
  Widget build(BuildContext context) {
    return _largeView();
  }

  Widget _largeView() {
    return Scaffold(
      appBar: AppBar(
        title: Text(title),
        backgroundColor: Theme.of(context).colorScheme.primaryContainer,
      ),
      body: SafeArea(
        child: Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Padding(
              padding: const EdgeInsets.all(8.0),
              child: SizedBox(
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
            ),
            const Gap(16),
            Padding(
              padding: const EdgeInsets.all(8.0),
              child: SizedBox(
                width: 600,
                child: widget.items[_selectedItemIndex].widget,
              ),
            ),
            const Gap(16),
            Flexible(
              child: BlocBuilder<SelectedItemBloc, (int, int)?>(
                  builder: (context, selecteState) {
                if (selecteState == null) {
                  return const Align(
                    alignment: Alignment.topCenter,
                    child: Text("Select a job to view logs"),
                  );
                }
                return BlocBuilder<JobQueueBloc, JobQueueState>(
                  builder: (context, state) {
                    return state.when(
                      initial: () => const SizedBox.shrink(),
                      update: (modelMap) {
                        final jobQueueItem = modelMap[selecteState.$1];
                        final jobModel = jobQueueItem.jobList[selecteState.$2];
                        return LogWidget(
                          log: jobModel.log,
                          started:
                              jobModel.jobStateStage == JobStateStage.started,
                        );
                      },
                    );
                  },
                );
              }),
            ),
          ],
        ),
      ),
    );
  }

  final String jobName =
      "BEC-AI-Chatbot__eoi-submission__book-3__proposed-solution";
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
