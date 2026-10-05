import json,os,re,struct,sys,tempfile,unicodedata,zipfile,zlib
from pathlib import Path
from dotenv import load_dotenv
try:import requests
except ImportError:requests=None
try:from docx2python import docx2python
except ImportError:sys.exit('Missing dependency. Install it with:\n    pip install docx2python --break-system-packages')
try:import docx as _pydocx
except ImportError:_pydocx=None
try:from PIL import Image as _PILImage
except ImportError:_PILImage=None
try:import imagecodecs as _imagecodecs
except ImportError:_imagecodecs=None
INPUT_PATH='C:\\Users\\Administrator\\Downloads\\DocQuesBuild\\test'
OUTPUT_PATH='C:\\Users\\Administrator\\Downloads\\DocQuesBuild\\test'
load_dotenv()
BEARER_TOKEN=os.getenv('BEARER_TOKEN')
SFM_UPLOAD_URL=os.getenv('SFM_UPLOAD_URL','https://campus.psgtech.ac.in/sfm/api/files/upload')
SFM_VIEW_URL=os.getenv('SFM_VIEW_URL','https://campus.psgtech.ac.in/sfm/api/files/viewfileurl')
SFM_UPLOAD_PATH=os.getenv('SFM_UPLOAD_PATH','/DOTECOE/atlas-9f3c7a1d')
MIN_UPLOAD_IMAGE_BYTES=int(os.getenv('MIN_UPLOAD_IMAGE_BYTES','2048'))
UPLOAD_SUPPORTED_EXTENSIONS={'.jpg','.jpeg','.png','.gif','.bmp','.webp'}
CATEGORY_HEADER_RE=re.compile('^[A-E]\\)?\\.?\\s*[A-Za-z][A-Za-z ]*SKILLS?\\s*\\(\\s*\\d+\\s*[Mm]arks?\\)?\\s*$')
STRAY_LETTER_RE=re.compile('^[A-E]\\)?\\.?\\s*$')
ROMAN_HEADING_RE=re.compile('^(VIII|III|VII|II|IV|VI|IX|I|V|X)\\b[).]?\\s*(.*)$')
SUBQ_RE=re.compile('^(\\d+)[).]\\s+(.*)$')
SUBJ_CODE_RE=re.compile('(?<!\\d)(?:[A-Za-z]{1,3}\\d{5,7}|\\d{9,10})(?!\\d)')
def extract_subj_code(stem:str)->str:
	m=SUBJ_CODE_RE.search(stem)
	if m:return m.group(0)
	return stem.split(' ',1)[0]
IMAGE_MARKER_RE=re.compile('----media/(image\\d+\\.\\w+)----')
IMAGE_ALT_MARKER_RE=re.compile('----Image alt text---->.*?<(?=----media/)',re.DOTALL)
IMAGE_ALT_MARKER_LOOSE_RE=re.compile('----Image alt text---->[^<\\n]*<?')
def strip_image_alt_markers(text:str)->str:text=IMAGE_ALT_MARKER_RE.sub('',text);return IMAGE_ALT_MARKER_LOOSE_RE.sub('',text)
LETTER_LABEL_RE=re.compile('^[A-Ea-e]\\)?\\.?$')
MARKS_TABLE_HEADER_RE=re.compile('^S\\.?\\s*NO\\.?$',re.IGNORECASE)
ROMAN_VALUES={'I':1,'II':2,'III':3,'IV':4,'V':5,'VI':6,'VII':7,'VIII':8,'IX':9,'X':10}
ROMAN_INT_TO_STR={v:k for(k,v)in ROMAN_VALUES.items()}
def _int_to_roman(n):return ROMAN_INT_TO_STR.get(n,str(n))
def clean_cell(paragraphs):
	if paragraphs is None:return''
	if isinstance(paragraphs,str):text=paragraphs
	else:text='\n'.join(str(p)for p in paragraphs if p)
	return strip_image_alt_markers(text.replace('\t',' ')).strip()
def _lineage_says_table(par):lineage=getattr(par,'lineage',None);return bool(lineage)and len(lineage)>1 and lineage[1]=='tbl'
def _is_real_table(pars_table):
	for row in pars_table:
		for cell in row:
			for par in cell:
				if _lineage_says_table(par):return True
	return False
def _flatten_entry(text_table):
	lines=[]
	for row in text_table:
		for cell in row:
			t=clean_cell(cell)
			if t:lines.append(t)
	return'\n'.join(lines)
def _flatten_table(rows):
	lines=[]
	for row in rows:
		cells=[clean_cell(cell)for cell in row];cells=[cell for cell in cells if cell]
		if cells:lines.append(' | '.join(cells))
	return'\n'.join(lines)
def extract_embedded_media(docx_path:Path,out_dir:Path):
	media_dir=out_dir/f"{docx_path.stem}_media";media_dir.mkdir(parents=True,exist_ok=True);extracted=[]
	try:
		with zipfile.ZipFile(docx_path)as archive:
			for name in archive.namelist():
				if not name.startswith('word/media/')or name.endswith('/'):continue
				target=media_dir/Path(name).name
				with archive.open(name)as src,open(target,'wb')as dst:dst.write(src.read())
				extracted.append(target.name)
	except Exception:return None,[]
	return media_dir.name,extracted
def _pad_jpeg(data:bytes,target_size:int)->bytes:
	needed=target_size-len(data)
	if needed<=0 or not data.endswith(b'\xff\xd9'):return data
	overhead=4;payload_len=max(needed-overhead,0);segments=b'';remaining=payload_len
	while remaining>0:chunk=min(remaining,65533);segments+=b'\xff\xfe'+struct.pack('>H',chunk+2)+b'\x00'*chunk;remaining-=chunk
	return data[:-2]+segments+data[-2:]
def _pad_png(data:bytes,target_size:int)->bytes:
	needed=target_size-len(data)
	if needed<=0 or not data.endswith(b'IEND\xaeB`\x82'):return data
	overhead=20;payload_len=max(needed-overhead,1);chunk_type=b'tEXt';chunk_data=b'Padding\x00'+b'0'*payload_len;chunk=struct.pack('>I',len(chunk_data))+chunk_type+chunk_data+struct.pack('>I',zlib.crc32(chunk_type+chunk_data)&4294967295);iend_start=len(data)-12;return data[:iend_start]+chunk+data[iend_start:]
def pad_image_to_min_size(path:Path,min_bytes:int=MIN_UPLOAD_IMAGE_BYTES)->bool:
	try:data=path.read_bytes()
	except Exception:return False
	if len(data)>=min_bytes:return True
	if data.startswith(b'\xff\xd8\xff'):padded=_pad_jpeg(data,min_bytes)
	elif data.startswith(b'\x89PNG\r\n\x1a\n'):padded=_pad_png(data,min_bytes)
	else:return False
	if len(padded)<min_bytes:return False
	path.write_bytes(padded);return True
def convert_unsupported_image_format(path:Path):
	ext=path.suffix.lower()
	if ext in UPLOAD_SUPPORTED_EXTENSIONS:return
	out_path=path.with_suffix('.png')
	def _save_array_as_png(arr):
		img=_PILImage.fromarray(arr)
		if img.mode not in('RGB','RGBA','L'):img=img.convert('RGB')
		img.save(out_path,'PNG');return out_path
	pil_error=None
	if _PILImage is not None:
		try:
			img=_PILImage.open(path);img.load()
			if img.mode not in('RGB','RGBA','L'):img=img.convert('RGB')
			img.save(out_path,'PNG');return out_path
		except Exception as e:pil_error=e
	if _imagecodecs is not None and _PILImage is not None:
		try:arr=_imagecodecs.imread(path.read_bytes());return _save_array_as_png(arr)
		except Exception as e:
			if ext in('.emf','.wmf','.svg'):print(f"    [WARN] {path.name} is a vector format ('{ext}') -- neither Pillow nor imagecodecs can rasterize vector drawings (they only decode pixel data); a rendering engine like LibreOffice or Inkscape would be needed to convert it -- skipping");return
			print(f"    [WARN] couldn't convert unsupported format {path.name} ({ext}): {e}");return
	if _PILImage is None:print(f"    [WARN] {path.name} has unsupported format '{ext}' and Pillow isn't installed to attempt a conversion -- install with:\n        pip install Pillow imagecodecs --break-system-packages")
	else:print(f"    [WARN] {path.name} has unsupported format '{ext}'; Pillow couldn't read it ({pil_error}) and 'imagecodecs' isn't installed to try further -- install with:\n        pip install imagecodecs --break-system-packages")
INLINE_IMAGE_FALLBACK=os.getenv('INLINE_IMAGE_FALLBACK','1')!='0'
INLINE_IMAGE_MAX_BYTES=int(os.getenv('INLINE_IMAGE_MAX_BYTES','1500000'))
_IMAGE_MIME={'.jpg':'image/jpeg','.jpeg':'image/jpeg','.png':'image/png','.gif':'image/gif','.bmp':'image/bmp','.webp':'image/webp'}
def local_image_as_data_uri(path:Path)->str:
	"""Fallback when no uploaded URL exists: embed the image itself so the <img> tag still renders."""
	import base64
	use=path
	if path.suffix.lower()not in _IMAGE_MIME:
		use=convert_unsupported_image_format(path)
		if use is None:return''
	try:data=use.read_bytes()
	except Exception:return''
	if len(data)>INLINE_IMAGE_MAX_BYTES:print(f"    [WARN] {use.name} is {len(data)} bytes, over INLINE_IMAGE_MAX_BYTES={INLINE_IMAGE_MAX_BYTES}; not inlined");return''
	return f"data:{_IMAGE_MIME.get(use.suffix.lower(),'image/png')};base64,"+base64.b64encode(data).decode('ascii')
def upload_images_to_s3(image_paths,docx_stem):
	if not image_paths or requests is None:return{}
	headers={}
	if BEARER_TOKEN:headers['Authorization']=f"Bearer {BEARER_TOKEN}"
	mapping={}
	for local_path in image_paths:
		upload_path=local_path
		if local_path.suffix.lower()not in UPLOAD_SUPPORTED_EXTENSIONS:
			converted_path=convert_unsupported_image_format(local_path)
			if converted_path is None:continue
			upload_path=converted_path
		if not pad_image_to_min_size(upload_path):print(f"    [WARN] {upload_path.name} is under {MIN_UPLOAD_IMAGE_BYTES} bytes and couldn't be padded (unrecognized format) -- the upload API will likely reject it")
		unique_name=f"{docx_stem}_{upload_path.name}";params={'Path':SFM_UPLOAD_PATH,'Filename':unique_name}
		try:
			with open(upload_path,'rb')as fh:resp=requests.post(SFM_UPLOAD_URL,params=params,files={'file':(unique_name,fh)},headers=headers,timeout=120)
			resp_body=None
			try:resp_body=resp.json()
			except ValueError:pass
			already_exists=resp.status_code==400 and isinstance(resp_body,dict)and'already exists'in str(resp_body.get('message','')).lower()
			if already_exists:key=f"{SFM_UPLOAD_PATH.strip("/")}/{unique_name}";print(f"    [REUSE] {upload_path.name} already exists in S3 at {key}, reusing it")
			elif not resp.ok:print(f"    [WARN] image upload failed for {upload_path.name}: {resp.status_code} -- {resp.text[:300]}");continue
			else:
				result=resp_body or{}
				if result.get('isError'):print(f"    [WARN] image upload reported error for {upload_path.name}: {result.get("message")}");continue
				key=(result.get('result')or{}).get('key')
				if not key:print(f"    [WARN] upload succeeded but no 'key' returned for {upload_path.name}");continue
		except Exception as e:print(f"    [WARN] image upload failed for {upload_path.name}: {e}");continue
		image_url=''
		try:
			view_resp=requests.get(SFM_VIEW_URL,params={'filepath':key},headers=headers,timeout=30)
			if view_resp.ok:
				view_result=view_resp.json()
				if not view_result.get('isError'):image_url=(view_result.get('result')or{}).get('url','')
		except Exception as e:print(f"    [WARN] could not fetch initial view URL for {upload_path.name}: {e}")
		mapping[local_path.name]={'filePath':key,'imageUrl':image_url}
	return mapping
_W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_SUPERSCRIPT_CHARS={'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','+':'⁺','-':'⁻','−':'⁻','=':'⁼','(':'⁽',')':'⁾','a':'ᵃ','b':'ᵇ','c':'ᶜ','d':'ᵈ','e':'ᵉ','f':'ᶠ','g':'ᵍ','h':'ʰ','i':'ⁱ','j':'ʲ','k':'ᵏ','l':'ˡ','m':'ᵐ','n':'ⁿ','o':'ᵒ','p':'ᵖ','r':'ʳ','s':'ˢ','t':'ᵗ','u':'ᵘ','v':'ᵛ','w':'ʷ','x':'ˣ','y':'ʸ','z':'ᶻ'}
_SUBSCRIPT_CHARS={'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','+':'₊','-':'₋','−':'₋','=':'₌','(':'₍',')':'₎','a':'ₐ','e':'ₑ','h':'ₕ','i':'ᵢ','j':'ⱼ','k':'ₖ','l':'ₗ','m':'ₘ','n':'ₙ','o':'ₒ','p':'ₚ','r':'ᵣ','s':'ₛ','t':'ₜ','u':'ᵤ','v':'ᵥ','x':'ₓ'}
_M_NS='http://schemas.openxmlformats.org/officeDocument/2006/math'
_XML_NS='http://www.w3.org/XML/1998/namespace'
MATH_INLINE=('$','$')
MATH_DISPLAY=('$$','$$')
CONVERT_TYPED_UNICODE_SCRIPTS=True
_UNI_TO_SUP={v:k for(k,v)in _SUPERSCRIPT_CHARS.items()}
_UNI_TO_SUB={v:k for(k,v)in _SUBSCRIPT_CHARS.items()}
_UNI_TO_SUP['⁻']='-'
_UNI_TO_SUB['₋']='-'
_UNI_SCRIPT_RE=re.compile('['+re.escape(''.join(set(_UNI_TO_SUP)|set(_UNI_TO_SUB)))+']+')
_GREEK={'α':'alpha','β':'beta','γ':'gamma','δ':'delta','ε':'varepsilon','ϵ':'epsilon','ζ':'zeta','η':'eta','θ':'theta','ϑ':'vartheta','ι':'iota','κ':'kappa','λ':'lambda','μ':'mu','\u00b5':'mu','ν':'nu','ξ':'xi','π':'pi','ϖ':'varpi','ρ':'rho','ϱ':'varrho','σ':'sigma','ς':'varsigma','τ':'tau','υ':'upsilon','φ':'varphi','ϕ':'phi','χ':'chi','ψ':'psi','ω':'omega','Γ':'Gamma','Δ':'Delta','\u2206':'Delta','Θ':'Theta','Λ':'Lambda','Ξ':'Xi','Π':'Pi','Σ':'Sigma','Υ':'Upsilon','Φ':'Phi','Ψ':'Psi','Ω':'Omega','\u2126':'Omega'}
_OPS={'→':'rightarrow','←':'leftarrow','↔':'leftrightarrow','⇒':'Rightarrow','⇐':'Leftarrow','⇔':'Leftrightarrow','⇌':'rightleftharpoons','⇋':'leftrightharpoons','↑':'uparrow','↓':'downarrow','⟶':'longrightarrow','⟵':'longleftarrow','⟷':'longleftrightarrow','⟹':'Longrightarrow','⟸':'Longleftarrow','⟺':'Longleftrightarrow','↦':'mapsto','×':'times','÷':'div','·':'cdot','⋅':'cdot','∙':'cdot','±':'pm','∓':'mp','∘':'circ','∞':'infty','≈':'approx','≠':'neq','≡':'equiv','≤':'leq','≥':'geq','≪':'ll','≫':'gg','∝':'propto','∼':'sim','≅':'cong','≃':'simeq','∈':'in','∉':'notin','⊂':'subset','⊆':'subseteq','⊃':'supset','⊇':'supseteq','∪':'cup','∩':'cap','∅':'emptyset','∀':'forall','∃':'exists','¬':'neg','∧':'wedge','∨':'vee','∂':'partial','∇':'nabla','∑':'sum','∏':'prod','∫':'int','∬':'iint','∭':'iiint','∮':'oint','…':'ldots','⋯':'cdots','⋮':'vdots','⋱':'ddots','ℏ':'hbar','ℓ':'ell','∠':'angle','⊥':'perp','∥':'parallel','△':'triangle','∴':'therefore','∵':'because','√':'surd','⊕':'oplus','⊗':'otimes','†':'dagger','‡':'ddagger'}
_MATH_SYMBOLS={k:'\\'+v+' 'for(k,v)in{**_GREEK,**_OPS}.items()}
_MATH_SYMBOLS.update({'Α':'A','Β':'B','Ε':'E','Ζ':'Z','Η':'H','Ι':'I','Κ':'K','Μ':'M','Ν':'N','Ο':'O','Ρ':'P','Τ':'T','Χ':'X','ℝ':'\\mathbb{R}','ℕ':'\\mathbb{N}','ℤ':'\\mathbb{Z}','ℚ':'\\mathbb{Q}','ℂ':'\\mathbb{C}','−':'-','–':'-','—':'-','′':"'",'″':"''",'°':'^{\\circ}','℃':'^{\\circ}\\mathrm{C}','\u212b':'\\text{\\AA}','ⅆ':'\\mathrm{d}','ⅇ':'\\mathrm{e}','ⅈ':'\\mathrm{i}'})
_MATH_ESC={'{':'\\{','}':'\\}','%':'\\%','#':'\\#','$':'\\$','_':'\\_','\\':'\\backslash ','~':'\\sim ','^':'\\wedge '}
_TEXT_ESC={'{':'\\{','}':'\\}','%':'\\%','#':'\\#','$':'\\$','_':'\\_','&':'\\&','\\':'\\textbackslash ','~':'\\textasciitilde ','^':'\\textasciicircum '}
_XARROWS={'→':'\\xrightarrow','⟶':'\\xrightarrow','←':'\\xleftarrow','⟵':'\\xleftarrow','⇒':'\\xRightarrow','⟹':'\\xRightarrow','⇐':'\\xLeftarrow','⟸':'\\xLeftarrow','↔':'\\xleftrightarrow','⟷':'\\xleftrightarrow','⇔':'\\xLeftrightarrow','⟺':'\\xLeftrightarrow','⇌':'\\xrightleftharpoons','⇋':'\\xleftrightharpoons'}
_NARY={'∑':'\\sum','∏':'\\prod','∐':'\\coprod','∫':'\\int','∬':'\\iint','∭':'\\iiint','∮':'\\oint','⋃':'\\bigcup','⋂':'\\bigcap','⋁':'\\bigvee','⋀':'\\bigwedge','⨁':'\\bigoplus','⨂':'\\bigotimes'}
_DELIMS={'{':'\\{','}':'\\}','‖':'\\|','⟨':'\\langle ','⟩':'\\rangle ','〈':'\\langle ','〉':'\\rangle ','⌊':'\\lfloor ','⌋':'\\rfloor ','⌈':'\\lceil ','⌉':'\\rceil '}
_ACCENTS={'\u0302':'hat','^':'hat','\u0303':'tilde','~':'tilde','\u0304':'bar','¯':'bar','\u0305':'overline','\u0307':'dot','˙':'dot','\u0308':'ddot','¨':'ddot','\u20db':'dddot','\u20d7':'vec','→':'vec','\u030c':'check','ˇ':'check','\u0301':'acute','´':'acute','\u0300':'grave','`':'grave','\u0306':'breve','˘':'breve'}
_KNOWN_FUNCS={'sin','cos','tan','cot','sec','csc','arcsin','arccos','arctan','sinh','cosh','tanh','coth','ln','log','lg','exp','lim','liminf','limsup','max','min','sup','inf','det','gcd','deg','dim','ker','arg','hom','Pr'}
_COMPLEX_RE=re.compile('\\\\(?:d?frac|binom|sqrt|sum|prod|coprod|int|iint|iiint|oint|big\\w+|begin|overbrace|underbrace)')
_SIMPLE_BASE_RE=re.compile('[A-Za-z0-9()\\[\\]]+|\\\\[A-Za-z]+|\\\\(?:mathrm|mathbf)\\{[A-Za-z0-9]+\\}')
def _ln(el):
	tag=el.tag
	if not isinstance(tag,str):return''
	if tag.startswith('{'+_M_NS+'}'):return'm:'+tag[len(_M_NS)+2:]
	if tag.startswith('{'+_W_NS+'}'):return'w:'+tag[len(_W_NS)+2:]
	return''
def _child(el,name):
	if el is None:return None
	for c in el:
		if _ln(c)==name:return c
	return None
def _mval(el):return None if el is None else el.get('{'+_M_NS+'}val')
def _is_on(el):return el is not None and(_mval(el)or'1').lower()in('1','on','true')
def _pr_val(el,pr_name,key,default=None):
	c=_child(_child(el,pr_name),key)
	if c is None:return default
	v=_mval(c);return default if v is None else v
def _run_chars(run):return''.join(c.text or''for c in run if _ln(c)in('m:t','w:t'))
def _only_text(el):
	if el is None:return None
	parts=[]
	for c in el:
		n=_ln(c)
		if not n or n.endswith('Pr'):continue
		if n!='m:r':return None
		parts.append(_run_chars(c))
	return''.join(parts).strip()
def _plain_math(text,arr=False):
	out=[];n=len(text)
	for(i,ch)in enumerate(text):
		if ch in _MATH_SYMBOLS:out.append(_MATH_SYMBOLS[ch])
		elif ch=='&':out.append('&'if arr else'\\&')
		elif ch in _MATH_ESC:out.append(_MATH_ESC[ch])
		elif ch=='\u2009':out.append('\\,')
		elif ch in'\u00a0\u2002\u2003':out.append('\\;')
		elif ch.isspace():
			prev=text[i-1]if i else'';nxt=text[i+1]if i+1<n else''
			if prev.isalpha()and nxt.isalpha():out.append('\\;')
		else:out.append(ch)
	return''.join(out)
def _text_mode(text,arr=False):
	out=[];buf=[]
	def flush():
		if buf:out.append('\\text{'+''.join(buf)+'}');buf.clear()
	for ch in text:
		if ch=='&'and arr:flush();out.append('&')
		elif ch in _MATH_SYMBOLS:flush();out.append(_MATH_SYMBOLS[ch])
		elif ch in _TEXT_ESC:buf.append(_TEXT_ESC[ch])
		else:buf.append(ch)
	flush();return''.join(out)
def _split_scripts(text):
	pos=0
	for m in _UNI_SCRIPT_RE.finditer(text):
		if m.start()>pos:yield(False,text[pos:m.start()])
		yield(True,m.group(0));pos=m.end()
	if pos<len(text):yield(False,text[pos:])
def _script_body(raw):
	raw=raw.strip()
	if re.fullmatch('[A-Za-z]{2,}',raw):return'\\mathrm{'+raw+'}'
	return _plain_math(raw)
def _script_pieces_latex(pieces):return''.join(f"{kind}{{{_script_body(text)}}}"for(kind,text)in pieces)
def _script_group_latex(chunk):
	pieces=[]
	for ch in chunk:
		kind,val=('^',_UNI_TO_SUP[ch])if ch in _UNI_TO_SUP else('_',_UNI_TO_SUB[ch])
		if pieces and pieces[-1][0]==kind:pieces[-1][1]+=val
		else:pieces.append([kind,val])
	return _script_pieces_latex(pieces)
def _math_run(el,arr=False):
	text=_run_chars(el)
	if not text:return''
	rpr=_child(el,'m:rPr');sty=scr=None;nor=False
	if rpr is not None:sty=_mval(_child(rpr,'m:sty'));scr=_mval(_child(rpr,'m:scr'));nor=_is_on(_child(rpr,'m:nor'))
	if nor:return _text_mode(text,arr)
	body=''.join(_script_group_latex(c)if is_s else _plain_math(c,arr)for(is_s,c)in _split_scripts(text))
	if not body.strip():return body
	font={'script':'mathcal','fraktur':'mathfrak','double-struck':'mathbb','sans-serif':'mathsf','monospace':'mathtt'}.get(scr or'')
	if font is None:font={'p':'mathrm','b':'mathbf','bi':'boldsymbol'}.get(sty or'')
	wr=_child(el,'w:rPr');b=_child(wr,'w:b')
	if b is not None and b.get('{'+_W_NS+'}val','1').lower()not in('0','false','off')and font in(None,'mathrm'):font='mathbf'
	return f"\\{font}{{{body}}}"if font else body
def _kids(el,arr=False):return''.join(_omml(c,arr)for c in el)
def _arg(el,name,arr=False):
	c=_child(el,name);return _kids(c,arr).strip()if c is not None else''
def _base(s):
	s=s.strip()
	if not s:return'{}'
	return s if _SIMPLE_BASE_RE.fullmatch(s)else'{'+s+'}'
def _extensible_arrow(el,arr=False):
	n=_ln(el)
	if n not in('m:limLow','m:limUpp'):return None
	base=_child(el,'m:e')
	if base is None:return None
	lim=_arg(el,'m:lim',arr);up,low=(lim,'')if n=='m:limUpp'else('',lim)
	glyph=_only_text(base)
	if glyph in _XARROWS:return _XARROWS[glyph],up,low
	nested=[c for c in base if _ln(c)in('m:limLow','m:limUpp')]
	if glyph is None and len(nested)==1:
		inner=_extensible_arrow(nested[0],arr)
		if inner:return inner[0],up or inner[1],low or inner[2]
	return None
def _omml(el,arr=False):
	if not isinstance(el.tag,str):return''
	n=_ln(el)
	if n.endswith('Pr')or n=='w:del':return''
	if n in('m:r','w:r'):return _math_run(el,arr)
	if n=='m:f':
		num=_arg(el,'m:num',arr);den=_arg(el,'m:den',arr);kind=_pr_val(el,'m:fPr','m:type','bar')
		if kind=='noBar':return f"\\binom{{{num}}}{{{den}}}"
		if kind in('skw','lin'):return f"{{{num}}}/{{{den}}}"
		return f"\\frac{{{num}}}{{{den}}}"
	if n=='m:sSub':return _base(_arg(el,'m:e',arr))+'_{'+_arg(el,'m:sub',arr)+'}'
	if n=='m:sSup':return _base(_arg(el,'m:e',arr))+'^{'+_arg(el,'m:sup',arr)+'}'
	if n=='m:sSubSup':
		out=_base(_arg(el,'m:e',arr))
		if not _is_on(_child(_child(el,'m:sSubSupPr'),'m:subHide')):out+='_{'+_arg(el,'m:sub',arr)+'}'
		if not _is_on(_child(_child(el,'m:sSubSupPr'),'m:supHide')):out+='^{'+_arg(el,'m:sup',arr)+'}'
		return out
	if n=='m:sPre':return'{}_{'+_arg(el,'m:sub',arr)+'}^{'+_arg(el,'m:sup',arr)+'}'+_base(_arg(el,'m:e',arr))
	if n=='m:rad':
		deg=_arg(el,'m:deg',arr);e=_arg(el,'m:e',arr)
		if deg and not _is_on(_child(_child(el,'m:radPr'),'m:degHide')):return f"\\sqrt[{deg}]{{{e}}}"
		return f"\\sqrt{{{e}}}"
	if n=='m:nary':
		ch=_pr_val(el,'m:naryPr','m:chr','∫');op=_NARY.get(ch,_plain_math(ch).strip());integral=ch in'∫∬∭∮∯∰'
		loc=_pr_val(el,'m:naryPr','m:limLoc','subSup'if integral else'undOvr')
		if loc=='undOvr':op+='\\limits'
		elif not integral:op+='\\nolimits'
		pr=_child(el,'m:naryPr');sub=_arg(el,'m:sub',arr);sup=_arg(el,'m:sup',arr)
		if sub and not _is_on(_child(pr,'m:subHide')):op+='_{'+sub+'}'
		if sup and not _is_on(_child(pr,'m:supHide')):op+='^{'+sup+'}'
		return op+' '+_arg(el,'m:e',arr)
	if n=='m:d':
		beg=_pr_val(el,'m:dPr','m:begChr','(');end=_pr_val(el,'m:dPr','m:endChr',')');sep=_pr_val(el,'m:dPr','m:sepChr','|')
		items=[_kids(e,arr).strip()for e in el if _ln(e)=='m:e'];b=_DELIMS.get(beg,beg);e_=_DELIMS.get(end,end)
		if _COMPLEX_RE.search(''.join(items)):
			joiner=' \\middle| 'if sep=='|'else _plain_math(sep)
			return'\\left'+(b or'.')+joiner.join(items)+'\\right'+(e_ or'.')
		return b+_plain_math(sep).join(items)+e_
	if n=='m:acc':
		ch=_pr_val(el,'m:accPr','m:chr','\u0302');e=_arg(el,'m:e',arr);cmd=_ACCENTS.get(ch,'hat')
		if not re.fullmatch('[A-Za-z0-9]|\\\\[A-Za-z]+',e):cmd={'hat':'widehat','tilde':'widetilde','bar':'overline'}.get(cmd,cmd)
		return f"\\{cmd}{{{e}}}"
	if n=='m:bar':
		e=_arg(el,'m:e',arr);return f"\\overline{{{e}}}"if _pr_val(el,'m:barPr','m:pos','bot')=='top'else f"\\underline{{{e}}}"
	if n=='m:groupChr':
		ch=_pr_val(el,'m:groupChrPr','m:chr','⏟');top=_pr_val(el,'m:groupChrPr','m:pos','bot')=='top';e=_arg(el,'m:e',arr)
		if not e:return _plain_math(ch)
		if ch=='⏞':return f"\\overbrace{{{e}}}"
		if ch=='⏟':return f"\\underbrace{{{e}}}"
		if ch=='→'and top:return f"\\overrightarrow{{{e}}}"
		if ch=='←'and top:return f"\\overleftarrow{{{e}}}"
		if ch=='↔'and top:return f"\\overleftrightarrow{{{e}}}"
		return(f"\\overset{{{_plain_math(ch)}}}{{{e}}}"if top else f"\\underset{{{_plain_math(ch)}}}{{{e}}}")
	if n in('m:limLow','m:limUpp'):
		arrow=_extensible_arrow(el,arr)
		if arrow:
			cmd,up,low=arrow
			return f"{cmd}[{low}]{{{up}}}"if low else f"{cmd}{{{up}}}"
		base=_child(el,'m:e');lim=_arg(el,'m:lim',arr);word=_only_text(base);e=_arg(el,'m:e',arr)
		if n=='m:limLow'and word in _KNOWN_FUNCS:return'\\'+word+'_{'+lim+'}'
		return(f"\\underset{{{lim}}}{{{e}}}"if n=='m:limLow'else f"\\overset{{{lim}}}{{{e}}}")
	if n=='m:func':
		fn=_child(el,'m:fName');raw=_only_text(fn)
		if raw in _KNOWN_FUNCS:name='\\'+raw
		elif raw and raw.isalpha():name='\\operatorname{'+raw+'}'
		else:name=_kids(fn,arr).strip()if fn is not None else''
		return name+' '+_arg(el,'m:e',arr)
	if n=='m:eqArr':return'\\begin{aligned}'+' \\\\ '.join(_kids(e,True).strip()for e in el if _ln(e)=='m:e')+'\\end{aligned}'
	if n=='m:m':
		rows=[' & '.join(_kids(e,arr).strip()for e in mr if _ln(e)=='m:e')for mr in el if _ln(mr)=='m:mr']
		return'\\begin{matrix}'+' \\\\ '.join(rows)+'\\end{matrix}'
	if n=='m:borderBox':return f"\\boxed{{{_arg(el,'m:e',arr)}}}"
	if n=='m:phant':
		e=_arg(el,'m:e',arr);return e if _pr_val(el,'m:phantPr','m:show','1')in('1','on','true')else f"\\phantom{{{e}}}"
	return _kids(el,arr)
def omml_to_latex(el):
	tex=_omml(el)
	tex=re.sub('(\\\\[A-Za-z]+) +(?=[}_^)\\],])',lambda m:m.group(1),tex)
	return re.sub(' {2,}',' ',tex).strip()
def _braces_balanced(s):
	depth=0;i=0
	while i<len(s):
		c=s[i]
		if c=='\\':i+=2;continue
		if c=='{':depth+=1
		elif c=='}':
			depth-=1
			if depth<0:return False
		i+=1
	return depth==0
def _equation_plain_text(el):return''.join(t.text or''for t in el.iter()if _ln(t)in('m:t','w:t'))
def _equation_to_inline_text(el):
	"""Return the final text (with $...$ / $$...$$ delimiters) that replaces one OMML equation."""
	display=_ln(el)=='m:oMathPara'
	try:
		if display:
			parts=[omml_to_latex(m)for m in el if _ln(m)=='m:oMath'];parts=[p for p in parts if p]
			tex=parts[0]if len(parts)==1 else'\\begin{gathered}'+' \\\\ '.join(parts)+'\\end{gathered}'if parts else''
		else:tex=omml_to_latex(el)
		if not _braces_balanced(tex):raise ValueError('unbalanced braces in generated LaTeX')
	except Exception as e:
		print(f"  [WARN] equation could not be converted to LaTeX ({e}); kept as plain text: {_equation_plain_text(el)[:60]}");return _equation_plain_text(el)
	if not tex:return''
	o,c=MATH_DISPLAY if display else MATH_INLINE
	return f"{o}{tex}{c}"
def _script_kind(run):
	if _ln(run)!='w:r':return None
	va=_child(_child(run,'w:rPr'),'w:vertAlign')
	if va is None:return None
	return{'superscript':'^','subscript':'_'}.get(va.get('{'+_W_NS+'}val'))
def _run_text(run):return''.join(t.text or''for t in run if _ln(t)=='w:t')
def _is_script_run(run):return bool(_script_kind(run))and bool(_run_text(run).strip())
def _rewrite_math_and_scripts(document_xml:bytes):
	from lxml import etree
	root=etree.fromstring(document_xml);changed=False;space='{'+_XML_NS+'}space'
	def swap_for_text(old,text):
		run=etree.Element('{'+_W_NS+'}r');t=etree.SubElement(run,'{'+_W_NS+'}t');t.text=text;t.set(space,'preserve');run.tail=old.tail;old.getparent().replace(old,run)
	for para in list(root.iter('{'+_M_NS+'}oMathPara')):
		if para.getparent()is not None:swap_for_text(para,_equation_to_inline_text(para));changed=True
	for eq in list(root.iter('{'+_M_NS+'}oMath')):
		if eq.getparent()is None or any(_ln(a)=='m:oMath'for a in eq.iterancestors()):continue
		swap_for_text(eq,_equation_to_inline_text(eq));changed=True
	for parent in list(root.iter()):
		kids=list(parent);i=0
		while i<len(kids):
			if not _is_script_run(kids[i]):i+=1;continue
			j=i;pieces=[]
			while j<len(kids)and _is_script_run(kids[j]):
				kind=_script_kind(kids[j]);txt=_run_text(kids[j])
				if pieces and pieces[-1][0]==kind:pieces[-1][1]+=txt
				else:pieces.append([kind,txt])
				j+=1
			first=kids[i];ts=[t for t in first if _ln(t)=='w:t'];ts[0].text=MATH_INLINE[0]+_script_pieces_latex(pieces)+MATH_INLINE[1];ts[0].set(space,'preserve')
			for extra in ts[1:]:first.remove(extra)
			rpr=_child(first,'w:rPr');rpr.remove(_child(rpr,'w:vertAlign'))
			for run in kids[i+1:j]:parent.remove(run)
			changed=True;i=j
	if CONVERT_TYPED_UNICODE_SCRIPTS:
		o,c=MATH_INLINE
		for t in root.iter('{'+_W_NS+'}t'):
			if t.text and _UNI_SCRIPT_RE.search(t.text):
				t.text=''.join(f"{o}{_script_group_latex(chunk)}{c}"if is_s else chunk for(is_s,chunk)in _split_scripts(t.text));t.set(space,'preserve');changed=True
	if not changed:return document_xml,False
	return etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True),True
def docx_with_latex_math(docx_path:Path):
	try:
		with zipfile.ZipFile(docx_path)as zin:
			new_xml,changed=_rewrite_math_and_scripts(zin.read('word/document.xml'))
			if not changed:return docx_path,False
			fd,tmp_name=tempfile.mkstemp(suffix='.docx');os.close(fd)
			with zipfile.ZipFile(tmp_name,'w',zipfile.ZIP_DEFLATED)as zout:
				for item in zin.infolist():data=new_xml if item.filename=='word/document.xml'else zin.read(item.filename);zout.writestr(item,data)
		return Path(tmp_name),True
	except Exception as e:print(f"  [WARN] Could not convert equations/sub/superscripts to LaTeX in {docx_path.name}: {e}");return docx_path,False
def get_body_entries(docx_path:Path):
	read_path,is_temp=docx_with_latex_math(docx_path)
	try:
		with docx2python(str(read_path))as doc:
			body=doc.body
			try:body_pars=doc.body_pars
			except AttributeError:body_pars=None
			entries=[]
			for(idx,text_table)in enumerate(body):
				if body_pars is not None:is_real=_is_real_table(body_pars[idx])
				else:is_real=not(len(text_table)==1 and len(text_table[0])==1)
				entries.append((is_real,text_table))
			return _split_trailing_headings(entries)
	finally:
		if is_temp:
			try:os.remove(read_path)
			except OSError:pass
def _split_trailing_headings(entries):
	expanded=[]
	for(is_real,rows)in entries:
		if is_real and len(rows)>1:
			last_row=rows[-1]
			if last_row and last_row[0]:
				cell0=clean_cell(last_row[0]).rstrip('.').strip();is_paired=len(last_row)>=2 and clean_cell(last_row[1]).strip()and len(clean_cell(last_row[1]).strip())>=15;is_bare=len(last_row)==1
				if cell0.isdigit()and cell0!='0'and(is_paired or is_bare):expanded.append((is_real,rows[:-1]));expanded.append((is_real,[last_row]));continue
		expanded.append((is_real,rows))
	return expanded
def _digit_heading_info(is_real,rows):
	if len(rows)<1 or len(rows[0])<1:return False,None,[]
	cell0=clean_cell(rows[0][0]).rstrip('.').strip()
	if not cell0.isdigit():return False,None,[]
	if cell0=='0':return False,None,[]
	if len(rows[0])>=2:
		for row in rows[1:]:
			if len(row)>=1 and clean_cell(row[0]).strip():return False,None,[]
		buffer_lines=[clean_cell(rows[0][1])]
		for extra_row in rows[1:]:
			if len(extra_row)>=2:
				t=clean_cell(extra_row[1])
				if t:buffer_lines.append(t)
		return True,cell0,buffer_lines
	if len(rows)==1 and len(rows[0])==1:return True,cell0,[]
	return False,None,[]
def _is_digit_heading(is_real,rows):is_heading,_,_=_digit_heading_info(is_real,rows);return is_heading
def _is_lettered_subanswer_table(rows):return len(rows)==1 and len(rows[0])==2 and bool(LETTER_LABEL_RE.match(clean_cell(rows[0][0]).strip()))and bool(clean_cell(rows[0][1]).strip())
def _is_stop_marker(rows):
	if not rows or not rows[0]:return False
	first_cell=clean_cell(rows[0][0]).strip()
	if first_cell.upper().startswith('ALLOCATION OF MARKS'):return True
	if MARKS_TABLE_HEADER_RE.match(first_cell):return True
	return False
def collect_digit_blocks(entries):
	blocks=[];i=0;n=len(entries)
	while i<n:
		is_real,rows=entries[i];is_heading,digit,initial_buffer=_digit_heading_info(is_real,rows)
		if is_heading:
			buffer_lines=list(initial_buffer);i+=1
			while i<n:
				is_real2,rows2=entries[i];is_heading2,_,_=_digit_heading_info(is_real2,rows2)
				if is_heading2:break
				if _is_stop_marker(rows2):i+=1;break
				if not is_real2:
					flat=_flatten_entry(rows2)
					if flat.strip()and flat.strip()!='0':buffer_lines.append(flat)
					i+=1
				elif _is_lettered_subanswer_table(rows2):label=clean_cell(rows2[0][0]).strip();desc=clean_cell(rows2[0][1]).strip();buffer_lines.append(f"{label} {desc}");i+=1
				else:
					flat_table=_flatten_table(rows2)
					if flat_table.strip():buffer_lines.append(flat_table)
					i+=1
			blocks.append((digit,'\n'.join(buffer_lines)))
		else:i+=1
	return blocks
def parse_digit_block(digit,raw_text,flags):
	lines=[l.rstrip()for l in raw_text.split('\n')];sections=[];preamble_lines=[];just_saw_category_header=False;seen_any_roman=False;next_position=1
	for line in lines:
		stripped=line.strip()
		if not stripped:continue
		if CATEGORY_HEADER_RE.match(stripped):just_saw_category_header=True;continue
		if STRAY_LETTER_RE.match(stripped):continue
		m=ROMAN_HEADING_RE.match(stripped)
		if m:
			token,rest=m.group(1),m.group(2);accept=False
			if token!='I':accept=True
			elif not seen_any_roman or just_saw_category_header:accept=True
			if accept:
				is_expected_reset=not seen_any_roman or just_saw_category_header;token_value=ROMAN_VALUES[token]
				if not is_expected_reset and token_value!=next_position:flags.append({'digit':digit,'line':stripped,'reason':f"Roman numeral '{token}' found where '{_int_to_roman(next_position)}' was expected next - possible typo in the source document's own numbering (not a parsing ambiguity). Verify the sequence is correct."})
				seen_any_roman=True;just_saw_category_header=False;next_position=token_value+1;sections.append({'roman':token,'lines':[rest]if rest else[]});continue
			else:flags.append({'digit':digit,'line':stripped,'reason':"Line starts with 'I' but is not the first Roman heading and doesn't follow a category header - treated as body text, not a new subsection. Verify this is correct."})
		else:
			am=SUBQ_RE.match(stripped)
			if am and just_saw_category_header:roman_token=_int_to_roman(next_position);seen_any_roman=True;just_saw_category_header=False;next_position+=1;sections.append({'roman':roman_token,'lines':[am.group(2)]if am.group(2)else[]});flags.append({'digit':digit,'line':stripped,'reason':f"Section heading used Arabic numeral '{am.group(1)}' instead of a Roman numeral in the source document - normalized to '{roman_token}' since it directly follows a category header. Verify this is correct."});continue
		just_saw_category_header=False
		if sections:sections[-1]['lines'].append(stripped)
		else:preamble_lines.append(stripped)
	if not sections:description='\n'.join(preamble_lines).strip();return{'Qno':digit,'Description':description}
	sub_divisions=[]
	for sec in sections:
		roman=sec['roman'];body_lines=sec['lines'];sub_questions=[];description_lines=[]
		for bl in body_lines:
			sm=SUBQ_RE.match(bl)
			if sm:sub_questions.append({'Qno':sm.group(1),'ParentQuestionId':roman,'Description':sm.group(2).strip()})
			elif not sub_questions:description_lines.append(bl)
			elif sub_questions:sub_questions[-1]['Description']+='\n'+bl
		node={'Qno':roman,'ParentQuestionId':digit,'Description':'\n'.join(description_lines).strip()}
		if sub_questions:node['SubQuestions']=sub_questions
		sub_divisions.append(node)
	return{'Qno':digit,'SubDivisions':sub_divisions}
def extract_questions_from_docx(docx_path:Path):entries=get_body_entries(docx_path);blocks=collect_digit_blocks(entries);flags=[];questions=[parse_digit_block(digit,raw_text,flags)for(digit,raw_text)in blocks];return questions,flags
def has_omml_equations(docx_path:Path)->bool:
	if _pydocx is None:return False
	try:d=_pydocx.Document(str(docx_path));return'<m:oMath'in d.element.xml
	except Exception:return False
def has_embedded_images(docx_path:Path)->bool:
	if _pydocx is None:return False
	try:d=_pydocx.Document(str(docx_path));xml=d.element.xml;return'<w:drawing'in xml or'<a:blip'in xml
	except Exception:return False
def substitute_image_placeholders(questions,url_by_filename):
	def repl(m):
		fname=m.group(1);url=url_by_filename.get(fname)
		if not url:return m.group(0)
		return f'<img src="{url}" alt="{fname}" />'
	def process_node(node):
		if node.get('Description')is not None:node['Description']=IMAGE_MARKER_RE.sub(repl,strip_image_alt_markers(node['Description']))
		for child in node.get('SubDivisions',[])or[]:process_node(child)
		for child in node.get('SubQuestions',[])or[]:process_node(child)
	for q in questions:process_node(q)
def load_existing_media_map(json_path:Path)->dict:
	if not json_path.exists():return{}
	try:
		with open(json_path,'r',encoding='utf-8')as f:data=json.load(f)
	except Exception:return{}
	mapping={}
	for entry in data.get('_embedded_media')or[]:
		fname=entry.get('fileName')
		if fname and entry.get('filePath'):mapping[fname]={'filePath':entry['filePath'],'imageUrl':entry.get('imageUrl','')}
	return mapping
def fetch_view_url(file_path_key:str)->str:
	if requests is None:return''
	headers={}
	if BEARER_TOKEN:headers['Authorization']=f"Bearer {BEARER_TOKEN}"
	try:
		resp=requests.get(SFM_VIEW_URL,params={'filepath':file_path_key},headers=headers,timeout=30)
		if resp.ok:
			result=resp.json()
			if not result.get('isError'):return(result.get('result')or{}).get('url','')
	except Exception:pass
	return''
def refresh_baked_urls(questions,url_by_basename:dict):
	if not url_by_basename:return
	patterns=[(re.compile('src="[^"]*'+re.escape(basename)+'[^"]*"'),f'src="{fresh_url}"')for(basename,fresh_url)in url_by_basename.items()]
	def process_node(node):
		if node.get('Description')is not None:
			text=node['Description']
			for(pattern,replacement)in patterns:text=pattern.sub(replacement,text)
			node['Description']=text
		for child in node.get('SubDivisions',[])or[]:process_node(child)
		for child in node.get('SubQuestions',[])or[]:process_node(child)
	for q in questions:process_node(q)
def refresh_json_file(json_path:Path)->bool:
	try:
		with open(json_path,'r',encoding='utf-8')as f:data=json.load(f)
	except Exception as e:print(f"  [ERROR] could not read {json_path.name}: {e}");return False
	media=data.get('_embedded_media')or[]
	if not media:print(f"  [SKIP] {json_path.name}: no embedded media to refresh");return False
	url_by_basename={};refreshed=0
	for entry in media:
		file_path=entry.get('filePath');fname=entry.get('fileName')
		if not file_path:continue
		fresh_url=fetch_view_url(file_path)
		if not fresh_url:print(f"    [WARN] could not refresh URL for {fname} ({file_path})");continue
		entry['imageUrl']=fresh_url;url_by_basename[Path(file_path).name]=fresh_url;refreshed+=1
	if not url_by_basename:print(f"  [SKIP] {json_path.name}: no URLs could be refreshed");return False
	refresh_baked_urls(data.get('questions',[]),url_by_basename)
	with open(json_path,'w',encoding='utf-8')as f:json.dump(data,f,ensure_ascii=False,indent=2)
	print(f"  [OK] {json_path.name}: refreshed {refreshed}/{len(media)} image URL(s)");return True
def process_file(docx_path:Path,out_dir:Path):
	has_equations=has_omml_equations(docx_path)
	try:questions,flags=extract_questions_from_docx(docx_path)
	except Exception as e:print(f"  [ERROR] Failed to process {docx_path.name}: {e}");return'error'
	if not questions:print(f"  [SKIP] {docx_path.name}: no questions found");return'no-questions'
	full_text=json.dumps(questions,ensure_ascii=False);media_dir_name,embedded_media=None,[]
	if has_embedded_images(docx_path)or IMAGE_MARKER_RE.search(full_text):media_dir_name,embedded_media=extract_embedded_media(docx_path,out_dir)
	uploaded_media=[]
	if embedded_media and media_dir_name:
		media_dir=out_dir/media_dir_name;existing_out_path=out_dir/(docx_path.stem+'.json');already_uploaded=load_existing_media_map(existing_out_path);url_map={};to_upload=[]
		for fname in embedded_media:
			if fname in already_uploaded:old=already_uploaded[fname];fresh_url=fetch_view_url(old['filePath'])or old['imageUrl'];url_map[fname]={'filePath':old['filePath'],'imageUrl':fresh_url};print(f"    [REUSE] {fname} already uploaded at {old["filePath"]}, skipping re-upload")
			else:to_upload.append(fname)
		if to_upload:image_paths=[media_dir/fname for fname in to_upload];new_results=upload_images_to_s3(image_paths,docx_path.stem);url_map.update(new_results)
		for fname in embedded_media:
			entry={'fileName':fname}
			if fname in url_map:entry.update(url_map[fname])
			uploaded_media.append(entry)
		url_by_filename={fname:info['imageUrl']for(fname,info)in url_map.items()if info.get('imageUrl')}
		if INLINE_IMAGE_FALLBACK:
			for fname in embedded_media:
				if fname in url_by_filename:continue
				data_uri=local_image_as_data_uri(media_dir/fname)
				if data_uri:url_by_filename[fname]=data_uri;print(f"    [INLINE] {fname}: no uploaded URL available, embedded as a data URI instead")
		if url_by_filename:substitute_image_placeholders(questions,url_by_filename)
	leftover=sorted(set(IMAGE_MARKER_RE.findall(json.dumps(questions,ensure_ascii=False))))
	if leftover:
		flags.append({'file':docx_path.name,'reason':f"image placeholder(s) could not be turned into <img> tags: {', '.join(leftover)} (upload failed and the local fallback was unavailable, e.g. unconvertible EMF/WMF)"})
		print(f"    [WARN] unresolved image placeholder(s) left in text: {', '.join(leftover)}")
	subj_code=extract_subj_code(docx_path.stem);output={'subj_code':subj_code,'questions':questions}
	if flags:output['_review_flags']=flags
	if has_equations:output['_review_flags']=output.get('_review_flags',[])+[{'file':docx_path.name,'reason':'document contains equation object(s); converted to LaTeX ($...$ inline, $$...$$ display) - spot-check the rendered output'}]
	if media_dir_name:output['_media_dir']=media_dir_name
	if uploaded_media:output['_embedded_media']=uploaded_media
	elif embedded_media:output['_embedded_media']=embedded_media
	out_path=out_dir/(docx_path.stem+'.json')
	with open(out_path,'w',encoding='utf-8')as f:json.dump(output,f,ensure_ascii=False,indent=2)
	flag_note=f", {len(flags)} line(s) flagged for review"if flags else'';print(f"  [OK] {docx_path.name} -> {out_path.name} ({len(questions)} question(s){flag_note})");return'ok'
def run_pipeline(input_path:Path,output_path:Path)->dict:
	if not input_path.exists():raise FileNotFoundError(f"Path does not exist: {input_path}")
	if input_path.is_file():
		if input_path.suffix.lower()!='.docx':raise ValueError(f"Not a .docx file: {input_path}")
		docx_files=[input_path]
	else:docx_files=sorted(input_path.rglob('*.docx'));docx_files=[f for f in docx_files if not f.name.startswith('~$')]
	if not docx_files:raise FileNotFoundError(f"No .docx files found at: {input_path}")
	seen_stems={}
	for f in docx_files:seen_stems.setdefault(f.stem,[]).append(f)
	dupes={stem:paths for(stem,paths)in seen_stems.items()if len(paths)>1}
	if dupes:
		print('[WARNING] Duplicate filenames found across subfolders - later files will overwrite earlier JSON output:')
		for(stem,paths)in dupes.items():
			for p in paths:print(f"    {stem}.docx  <-  {p}")
	print(f"Found {len(docx_files)} .docx file(s).");output_path.mkdir(parents=True,exist_ok=True);ok_count=0;no_q_count=0;image_count=0;eqn_count=0;error_count=0
	for docx_file in docx_files:
		result=process_file(docx_file,output_path)
		if result=='ok':ok_count+=1
		elif result=='has-image':image_count+=1
		elif result=='has-equations':eqn_count+=1
		elif result=='no-questions':no_q_count+=1
		else:error_count+=1
	print(f"\nDone. {ok_count}/{len(docx_files)} file(s) converted successfully.")
	if image_count:print(f"  {image_count} file(s) skipped: contain embedded image(s) (not handled yet)")
	if eqn_count:print(f"  {eqn_count} file(s) skipped: contain embedded equation object(s) (not handled yet)")
	if no_q_count:print(f"  {no_q_count} file(s) skipped: no questions found (check structure)")
	if error_count:print(f"  {error_count} file(s) skipped: processing error")
	return{'total':len(docx_files),'ok':ok_count,'no_questions':no_q_count,'images':image_count,'equations':eqn_count,'errors':error_count,'duplicates':list(dupes.items())}
def main():
	args=sys.argv[1:]
	if args and args[0]=='--refresh':
		refresh_args=args[1:];target=Path(refresh_args[0]).expanduser().resolve()if refresh_args else Path(OUTPUT_PATH).expanduser().resolve()
		if not target.exists():sys.exit(f"Path does not exist: {target}")
		json_files=[target]if target.is_file()else sorted(target.glob('*.json'))
		if not json_files:sys.exit(f"No .json file(s) found at: {target}")
		print(f"Refreshing image URLs in {len(json_files)} file(s)...");ok_refresh=sum(1 for jf in json_files if refresh_json_file(jf));print(f"\nDone. {ok_refresh}/{len(json_files)} file(s) refreshed.");return
	input_path=Path(args[0]).expanduser().resolve()if len(args)>=1 else Path(INPUT_PATH).expanduser().resolve();output_path=Path(args[1]).expanduser().resolve()if len(args)>=2 else Path(OUTPUT_PATH).expanduser().resolve()
	try:run_pipeline(input_path,output_path)
	except(FileNotFoundError,ValueError)as e:sys.exit(str(e))
if __name__=='__main__':main()