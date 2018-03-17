
# coding: utf-8
from WindPy import *

from events_mointor.gdzjc_event import gdzjc_event_main
from events_mointor.ggcgbd_event import ggcgbd_event_main
from events_mointor.ggzjc_event import ggzjc_event_main
from events_mointor.hsgtzjl_event import hsgtzjl_event_main
from events_mointor.rzrq_event import rzrq_event_main
from events_mointor.zlzjl_event import zlzjl_event_main

from events_mointor.announcement_event import announcement_main
from email_sender.email_sender import EmailEngine


def event_engine_main():

    # announcement_main()

    w.start()
    gdzjc_event_main()
    ggzjc_event_main()
    zlzjl_event_main()

    ggcgbd_event_main()
    hsgtzjl_event_main()
    rzrq_event_main()


if __name__ == "__main__":
    event_engine_main()
    # zip_files()
    # send_email()