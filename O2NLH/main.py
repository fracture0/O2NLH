import argparse
from gooey import Gooey, GooeyParser
from readcsv import converter


@Gooey(program_name="Convertisseur Ventes O2", language="french", show_success_modal=False)
def main():
    parser = GooeyParser()
    parser.add_argument(
        "--ventes", help="Chemin du fichier à convertir", widget="FileChooser"
    )
    parser.add_argument(
        "--journal", help="Code journal", default="VE"
    )
    args = parser.parse_args()
    print("fichier ventes : {}".format(args.ventes), flush=True)
    print("code journal : {}".format(args.journal), flush=True)
    nblignes = converter(args.ventes, args.journal)
    print("nombre de lignes : {}".format(nblignes), flush=True)

if __name__ == '__main__':    
    main()
