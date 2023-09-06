import os
import shutil
import lxml
import re
from lxml import etree
from zipfile import ZipFile
from copy import deepcopy

def copyEle(ele):
    res = etree.Element(ele.tag)
    res.text = ele.text
    for attr in ele.attrib: 
        res.attrib[attr] = ele.attrib[attr]
    return res

class DocHandle:
    def __init__(self, doc_path = 'test.docx', save_path='.'):
        self.docx_path = doc_path
        self.save_path = save_path
        self.docx_zip = ZipFile(self.docx_path)
        self.docx_zip.extractall(self.docx_path[:-5])
        self.pic_map = self.get_pic_map(self.docx_path[:-5]+'/word/_rels/document.xml.rels')
        self.documnet_xml = self.docx_path[:-5] + '/word/document.xml'
        self.tree = etree.parse(self.documnet_xml)
        self.doc_root = self.tree.getroot()
        self.doc_body = self.doc_root[0]
        self.index = {}
        self.iter_index = 0
        self.iter = self.get_iter()
        self.init_etree()
        self.media_list = [] 
        self.catalog = self.get_catalog()
        print('---------Got catalogue---------')
        print(self.catalog)
        print('--------------------------------')
        print('---------Begin to split---------------')
        self.get_context_index()
        
        
    def get_iter(self):
        for i,ele in enumerate(self.doc_body):
            self.iter_index = i
            # 开始标记
            yield (ele,ele)
            for child in ele.iter():
                yield (ele,child)
            # 结束标记
            yield(ele, None)
    
    def get_pic_map(self, path):
        # get picture relations from document.xml.rels
        res = {}
        with open(path, 'r') as f:
            for l in f:
                for line in l.split('<'):
                    m = re.search('Id=".*?"', line)
                    if m :
                        span = m.span()
                        key = line[span[0]:span[1]].lstrip('Id="').rstrip('"')
                    m = re.search('Target=".*?"', line)
                    if m :
                        span = m.span()
                        value = line[span[0]:span[1]].lstrip('Target="').rstrip('"')
                        res[key] = value
        return res

    def get_catalog_title(self, ele):
        # 获取目录元素内的标题序号、标题内容
        title = ''
        title_all = ''
        catalog_flag = False
        for i in ele.iter():
            if i.tag.endswith('instrText'):
                catalog_flag = True
            else:
                if i.text and i.text.strip():
                    title_all += i.text.strip()
                    if not i.text.isnumeric() and '.' not in i.text:
                        title += i.text.strip()
        title_raw = title_all.replace(title,'|||').split('|||')
        if len(title_raw)>1:
            title_id = title_raw[0]
        else:
            title_id = None
        
        if catalog_flag:
            return (title, title_id)
        else:
            return (None, None)
    
    def get_anchor(self, ele):
        for attr in ele.attrib:
            if attr.endswith('anchor'):
                return ele.attrib[attr]    

    def get_embed(self, ele):
        for attr in ele.attrib:
            if attr.endswith('embed'):
                return ele.attrib[attr]
    
    def get_catalog(self):
        res = []
        catalogBegin = False
        catalogEnd = False
        catalog_flag = False
        for (ele,child) in self.iter:
            # 判断大元素中是否有目录标记 
            # 大元素结束
            if child==None:
                # 若已经处于目录结构中且上一个大元素没有目录标记则认为目录结束
                if  catalogBegin and not catalog_flag:
                    catalogEnd = True
                if catalogBegin and catalogEnd:
                    return res
                self.append_ele(ele)
                continue
           
            # 大元素开始
            elif child == ele:
                # 默认无标记
                catalog_flag = False

                
            if child.tag.endswith('hyperlink'):
                title, title_id = self.get_catalog_title(ele)
                anchor = self.get_anchor(child)
                if title:
                    res.append((title_id, title, anchor))
                    catalog_flag = True
                    if not catalogBegin:
                        catalogBegin = True
            
            if child.tag.endswith('blip'):
                embed = self.get_embed(child)
                self.media_list.append(self.pic_map[embed])

        print('无目录或目录无超链接')
        return res
        
        
    def find_match_index(self, name, index):
        for i,(title_id,title,anchor) in enumerate(self.catalog):
            if anchor.strip() == name.strip():
                if i == 0:
                    return 0
                self.index[str(title_id)+str(title)] = index
                if self.big_title(self.catalog[i-1][0]):
                    return None
                else:
                    return i
        return None
                
    
    def get_name(self, ele):
        for attr in ele.attrib:
            if attr.endswith('name'):
                return ele.attrib[attr]
    
    def get_context_index(self):
        for i,(ele,child) in enumerate(self.iter):
            # 大元素结束
            if child==None:
                self.append_ele(ele)
                continue
            if child.tag.endswith('bookmarkStart'):
                name = self.get_name(child)
                catalog_index = self.find_match_index(name, self.iter_index)
                if catalog_index != None:
                    if catalog_index == 0:
                        self.save('0.0catalogue.docx')
                    else:
                        self.save(self.catalog[catalog_index-1][0]+self.catalog[catalog_index-1][1]+'.docx')
            if child.tag.endswith('blip'):
                embed = self.get_embed(child)
                self.media_list.append(self.pic_map[embed])
        if len(self.catalog)>0:
            self.save(self.catalog[-1][0]+self.catalog[-1][1]+'.docx')
                        
    def big_title(self, find_id):
        # 判断是否是大标题
        id_len = len(find_id.split('.'))
        for i,(title_id,title,anchor) in enumerate(self.catalog):
            if find_id == title_id:

                if i == len(self.catalog)-1:
                    return False
                next_part = self.catalog[i+1]
                if id_len >= len(next_part[0].split('.')):
                    return False
                elif id_len < len(next_part[0].split('.')):
                    return True
    
    def init_etree(self):
        # 初始化要保存的新word的document.xml
        self.new_root = copyEle(self.doc_root)
        self.etree = etree.ElementTree(self.new_root)
        self.new_body = copyEle(self.doc_root[0])
        self.new_root.append(self.new_body)
        return self.etree        
        
    def append_ele(self, ele):
        # 将element写入自身的etree缓存
        self.new_body.append(deepcopy(ele))

        
    def save(self, name):
        # 将自身的etree写入对应路径
        print(name)
        if not os.path.exists(self.save_path):
            os.makedirs(self.save_path)
        name = name.replace('/','')
        path = os.path.join(self.save_path, name)
        shutil.copy(self.docx_path, path)
        docx_zip = ZipFile(self.docx_path)
        docx_zip.extractall(path[:-5])
        documnet_xml = path[:-5] + '/word/document.xml'
        self.etree.write(documnet_xml, encoding='utf-8', xml_declaration=True, standalone=True)
        new_docx_file = ZipFile(path, mode='w')
        for i in self.docx_zip.namelist():
            if i.startswith('word/media/image') and  i.lstrip('word/') not in self.media_list:
                pass
            else:
                new_docx_file.write(path[:-5]+'/'+i,i)
        new_docx_file.close()
        docx_zip.close()
        try:
            shutil.rmtree(path[:-5])
        except:
            print(path[:-5])
        # 清空etree的缓存
        self.init_etree()
        self.media_list = []
        
    def close(self):
        self.docx_zip.close()
        shutil.rmtree(self.docx_path[:-5])
        
def word2md(word_path, md_path, pic_path = None):
    cmd_code = 'pandoc -f docx -t markdown_mmd -o %(md)s -s %(doc)s --extract-media=%(pic)s'
    if pic_path == None:
        pic = os.path.join(md_path,'pic')
    if not os.path.exists(md_path):
        os.makedirs(md_path)
    docx_list = os.listdir(word_path)
    for d in docx_list:
        md_name = d[:-5]+'.md'
        md = os.path.join(md_path, md_name)
        doc_p = os.path.join(word_path, d)
        cmd_c = cmd_code % {'md':md,'doc':doc_p,'pic':pic}
        os.popen(cmd_c)
    # shutil.rmtree(word_path)
        
if __name__ =="__main__":
    temp_file = './temp'
    word_name = 'test.docx'
    doc = DocHandle(doc_path = word_name, save_path=temp_file)
    doc.close()
    # word2md(temp_file, './md')

