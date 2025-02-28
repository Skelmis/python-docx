
.. _text_api:

Text-related objects
====================


|Paragraph| objects
-------------------

.. autoclass:: skelmis.docx.text.paragraph.Paragraph()
   :members:


|ParagraphFormat| objects
-------------------------

.. autoclass:: skelmis.docx.text.parfmt.ParagraphFormat()
   :members:


|Hyperlink| objects
-------------------

.. autoclass:: skelmis.docx.text.hyperlink.Hyperlink()
   :members:


|Run| objects
-------------

.. autoclass:: skelmis.docx.text.run.Run()
   :members:


|Font| objects
--------------

.. autoclass:: skelmis.docx.text.run.Font()
   :members:


|RenderedPageBreak| objects
---------------------------

.. autoclass:: skelmis.docx.text.pagebreak.RenderedPageBreak()
   :members:


|TabStop| objects
-----------------

.. autoclass:: skelmis.docx.text.tabstops.TabStop()
   :members:


|TabStops| objects
------------------

.. autoclass:: skelmis.docx.text.tabstops.TabStops()
   :members: clear_all

   .. automethod:: skelmis.docx.text.tabstops.TabStops.add_tab_stop(position, alignment=WD_TAB_ALIGNMENT.LEFT, leader=WD_TAB_LEADER.SPACES)
