PortFill word lists
===================
Every file here (one per language) tells PortFill what a label on a form means.
You can add your own words without a new program version: make a text file, for example
   my_words.txt
and put it in the folder   .portfill/words   inside your home folder (create it if it does not exist).
PortFill reads it at start-up, together with the files in this folder.

Format:
   [section]
   concept = word or phrase, another phrase, ...
Lines starting with # are comments. Capitals, accents and punctuation do not matter.

Sections: person (crew/passenger columns), ports (port-call columns), waste (waste columns),
          fields (single fields such as ship name), store_classes (kinds of ship's stores),
          stores (columns of the stores list: article, quantity, unit, location).

Rules of thumb
 * A one-word entry only matches when the whole label is exactly that word.
 * A phrase matches when all its words are in the label.
 * Words that are not 100% sure are better left out: an empty yellow field is safer than a wrong one.
 * Entries marked  # CHECK  in these files were translated by an AI and should be confirmed by a native speaker.
