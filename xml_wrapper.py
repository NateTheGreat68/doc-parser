from typing import Optional
import re
from xml.dom import minidom


class XMLWrapper:
    """
    This class adds some functionality to the existing xml.dom.minidom class.

    Init parameters:
      - xml_contents: A str containing the XML to be parsed.
      - text_node_tag = 'w:t': Optional. The XML tag name for nodes expected to
        contain text.
    """

    def __init__(
            self,
            xml_contents: str,
            text_node_tag: Optional[str] = 'w:t',
            ):
        self.minidom = minidom.parseString(xml_contents)
        self.text_node_tag = text_node_tag
        self.reset_to_root()

    def reset_to_root(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Resets the current target node to the document's root.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.minidom
            return self
        except:
            return None

    def set_node(
            self,
            node: minidom.Node,
            ) -> 'XMLWrapper':
        """
        Sets the target node.

        Parameters:
          - node: The node to be set.
        Return: This object, ready to call additional methods.
        """
        self.node = node

    def set_firstChild(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's first child.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.firstChild
            return self
        except:
            return None

    def set_lastChild(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's last child.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.lastChild
            return self
        except:
            return None

    def set_childNode(
            self,
            nth: int,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's nth child.

        Parameters:
          - nth: The index of the child node to set.
        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.childNodes[nth]
            return self
        except:
            return None

    def set_parentNode(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's parent.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.parentNode
            return self
        except:
            return None

    def set_nextSibling(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's next sibling.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.nextSibling
            return self
        except:
            return None

    def set_previousSibling(
            self,
            ) -> Optional['XMLWrapper']:
        """
        Sets the target node to the current node's previous sibling.

        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            self.node = self.node.previousSibling
            return self
        except:
            return None

    def set_text_node_by_pattern(
            self,
            pattern: str,
            ) -> Optional['XMLWrapper']:
        """
        Finds the first XML text node whose value matches the supplied regex pattern,
        and sets it as the current reference node within the wrapper.

        Parameters:
          - pattern: The regex pattern to search for.
        Return: This object, ready to call additional methods; None on failure.
        """
        r = re.compile(pattern)
        try:
            for n in self.node.getElementsByTagName(self.text_node_tag):
                if n.firstChild.nodeType == n.TEXT_NODE \
                        and r.search(n.firstChild.nodeValue):
                            self.node = n
                            return self
        except:
            return None

    def set_ancestor_node_by_tag(
            self,
            tag_name: str,
            ) -> Optional['XMLWrapper']:
        """
        Traverses up the XML DOM from the current node in search of the first
        ancestor which matches the supplied tag name.

        Parameters:
          - tag_name: The XML tag name to search for.
        Return: This object, ready to call additional methods; None on failure.
        """
        try:
            n = self.node
            while n.parentNode is not None:
                n = n.parentNode
                if n.tagName == tag_name:
                    self.node = n
                    return self
            return None
        except:
            return None

    def get_text_value(
            self,
            start_node: minidom.Node = None,
            ) -> Optional[str]:
        """
        Returns a compilation of all text values found within the XML DOM.

        Parameters:
          - start_node: The node to start at; used for recursion.
        Return: The found string; None on failure.
        """
        try:
            found_text = ''
            if not start_node:
                start_node = self.node
            for n in start_node.childNodes:
                if n.nodeType == n.TEXT_NODE: #Node is text
                    found_text += n.nodeValue
                elif n.tagName in self.special_tags: #Tag has special meaning
                    found_text += self.special_tags[n.tagName]
                else: #Descend into the child node.
                    found_text += self.get_text_value(n)
            return found_text
        except:
            return None
