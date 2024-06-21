import 'package:flutter/material.dart';
import 'package:flutter_bloc/flutter_bloc.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:ui/core/api_module.dart';
import 'package:ui/core/dependency.dart';
import 'package:ui/feature/api_doc/api_doc_widget.dart';
import 'package:ui/feature/faq/domain/presentation/faq_widget.dart';
import 'package:ui/feature/gseet_input/presentation/home_widget.dart';
import 'package:ui/feature/gsheet_template_gallery/gallery_widget.dart';
import 'package:ui/feature/screen/landing_screen.dart';
import 'package:ui/feature/shared/presentation/bloc/job_queue_bloc.dart';
import 'package:ui/feature/shared/presentation/bloc/selected_item_bloc.dart';

class MyApp extends StatelessWidget {
  const MyApp({super.key});
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Doculaboration',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
        textTheme: GoogleFonts.outfitTextTheme(Theme.of(context).textTheme),
      ),
      home: MultiBlocProvider(
        providers: [
          BlocProvider<JobQueueBloc>(create: (_) => getIt<JobQueueBloc>()),
          BlocProvider<SelectedItemBloc>(
            create: (_) => getIt<SelectedItemBloc>(),
          ),
        ],
        child: Builder(builder: (context) {
          return LandingScreen(
            initialSelectedItemIndex: 0,
            items: [
              DrawerItem(
                title: "Home",
                widget: HomeWidget(
                  jobListBloc: context.read<JobQueueBloc>(),
                  selectedItembloc: context.read<SelectedItemBloc>(),
                ),
                icon: const Icon(Icons.home),
              ),
              DrawerItem(
                title: "FAQ",
                widget: const FAQWidget(),
                icon: const Icon(Icons.question_mark),
              ),
              DrawerItem(
                title: "Templates",
                widget: GalleryWidget(
                  items: [
                    GalleryItem(
                      title: "Product Data Sheet(PDS)",
                      url:
                          "https://docs.google.com/spreadsheets/d/1efx8cDCGfM9xymswn8bOYmF4iZhtjyL2ShMoqMojL7Q/edit?usp=sharing",
                    ),
                    GalleryItem(
                      title: "Resume",
                      url:
                          "https://docs.google.com/spreadsheets/d/128SCmFDBY5lD1GOuKKCkuoG8g9jWZeTivhpcBMjTQzE/edit?usp=sharing",
                    ),
                  ],
                ),
                icon: const Icon(Icons.library_books),
              ),
              DrawerItem(
                title: "API Document",
                widget: ApiDocWidget(
                  url: "${getIt.get(
                    instanceName: kBaseUrl,
                    type: String,
                  )}/docs",
                ),
                icon: const Icon(Icons.api),
              ),
            ],
          );
        }),
      ),
    );
  }
}
