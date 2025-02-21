"""
# Version info

This script shows the version information of TkEasyGUI and system information.
Also Fig. is displayed using column.
"""

import subprocess
import TkEasyGUI as eg

WEB_SITE = "https://github.com/kujirahand/tkeasygui-python/"
S_COPY = eg.get_text("Copy")
S_CLOSE = eg.get_text("Close")



def help_window(script_name="test"):
    # define layout
    # col1 - layout
    layout1 = [
            [eg.Text(f"JROD-gui Project: {script_name}", font=("Arial", 14, "bold"), color="darkgreen")],
            [eg.Text(f"{eg.__doc__.strip()}", color="navy")],
            [eg.Frame("Version", expand_x=True,
                    layout=[
                        [   eg.Text("TkEasyGUI: "),
                            eg.Button(f"v{eg.__version__}", key="-b1-"),eg.Text(" "),
                            eg.Button("Web"),  ],
                    ],  ), ],
                ]
    #layout2 = [[eg.Image(size=(99, 100), filename="Leaf2.png")],]
    layout2 = [[eg.Frame("", layout = [
                    [eg.Image(size=(99, 100), filename="Leaf2.png")],])],
              ]      
    col1 = eg.Column(layout1, key="col1", expand_x=True)
    col2 = eg.Column(layout2, key="col2")
    
  
    layout = [ [col1, col2],
        #[main_lay,],
        #[eg.Text(f"JROD-gui Project: {script_name}", font=("Arial", 14, "bold"), color="darkgreen"),],
        #[eg.Text(f"{eg.__doc__.strip()}", color="navy"),],
        #[eg.Frame("Version",
        #        expand_x=True,
        #        layout=[
        #            [   eg.Text("TkEasyGUI: "),
        #                eg.Button(f"v{eg.__version__}", key="-b1-"),eg.Text(" "),
        #                eg.Button("Web"),  ],
        #        ],  )
        #],
        [eg.Frame("System info:",
                layout=[
                    [   eg.Multiline( f"{eg.get_system_info()}",
                            key="-sys-info-",
                            size=(60, 8),
                            expand_x=True,
                        )
                    ],
                    [eg.Button(S_COPY), eg.Button("Copy as Markdown")],
                ],
            )
        ],
        [eg.Column(
                layout=[[eg.Button("OK"), eg.Button(S_CLOSE)]],
                text_align="right",
                expand_x=True,
            ),
        ],
    ]
    # window create
    window = eg.Window("Version info", layout=layout, font=("Arial", 14 if eg.is_mac() else 12), 
                    row_padding=3, modal = True)
    # event loop
    for event, values in window.event_iter():
        print(f"# event = {event}, values = {values}")
        if event == "OK":
            #eg.popup("Thank you.")
            break
        if event == S_CLOSE:
            break
        if event in ["-b1-", "-b2-", "-b3-"]:
            btn: eg.Button = window[event]
            label = btn.get_text()
            eg.set_clipboard(label)
            eg.popup(f"Copied to clipboard:\n{label}")
        if event == S_COPY:
            text = window["-sys-info-"].get_text()
            eg.set_clipboard(text)
            eg.popup("Copied to clipboard.")
        if event == "Copy as Markdown":
            text = window["-sys-info-"].get_text()
            text = f"```\n{text}\n```\n"
            eg.set_clipboard(text)
            eg.popup("Copied markdown to clipboard.")
        if event == "Web":
            if eg.is_mac():
                subprocess.call(f"open {WEB_SITE}", shell=True)
            else:
                subprocess.call(f"start {WEB_SITE}", shell=True)
    return window

# -------------
if __name__ == "__main__":
    print("このスクリプトは直接実行されました")

    win = help_window("test")
    win.close()

