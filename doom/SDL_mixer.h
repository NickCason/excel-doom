/* Empty on purpose. The engine's i_sound.c includes <SDL_mixer.h> whenever FEATURE_SOUND is
   defined, but it uses nothing from it; the Excel build supplies its own sound module in
   dg_xl.c and has no SDL. */
