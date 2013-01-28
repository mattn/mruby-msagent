#include <windows.h>
#include <objbase.h>
#include <ctype.h>

#include <mruby.h>
#include <mruby/string.h>
#include <mruby/data.h>
#include <mruby/variable.h>
#include <stdio.h>

static IID IID_IAgentEx = {0x48D12BA0,0x5B77,0x11d1,{0x9E,0xC1,0x00,0xC0,0x4F,0xD7,0x08,0x1F}};
static CLSID CLSID_AgentServer = {0xD45FD2FC,0x5C6E,0x11D1,{0x9E,0xC1,0x00,0xC0,0x4F,0xD7,0x08,0x1F}};

typedef struct {
  IDispatch* pAgentEx;
  IDispatch* pCharacterEx;
  long request_id;
  mrb_value instance;
  mrb_state* mrb;
} mrb_msagent;

static void
msagent_free(mrb_state *mrb, void *p) {
  mrb_msagent* agent = (mrb_msagent*) p;
  if (agent->pCharacterEx) {
    agent->pCharacterEx->lpVtbl->Release(agent->pCharacterEx);
  }
  if (agent->pAgentEx) {
    agent->pAgentEx->lpVtbl->Release(agent->pAgentEx);
  }
  free(agent);
}

static const struct mrb_data_type msagent_type = {
  "msagent", msagent_free,
};

static BSTR
utf8_to_bstr(const char* str) {
  DWORD size = MultiByteToWideChar(CP_UTF8, 0, str, -1, NULL, 0);
  BSTR bs;
  if (size == 0) return NULL;
  bs = SysAllocStringLen(NULL, size - 1);
  MultiByteToWideChar(CP_UTF8, 0, str, -1, bs, size);
  return bs;
}

static mrb_value
emsg(mrb_state* mrb, HRESULT hr) {
  LPVOID buf = NULL;
  FormatMessageA(
    FORMAT_MESSAGE_ALLOCATE_BUFFER |
    FORMAT_MESSAGE_FROM_SYSTEM,
    NULL,
    hr,
    MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US),
    (LPSTR) &buf,
    0,
    NULL);
  if (buf != NULL) {
    mrb_value str = mrb_str_new_cstr(mrb, (char*) buf);
    LocalFree(buf);
    return str;
  }
  return mrb_str_new_cstr(mrb, "");
}

static mrb_value
mrb_msagent_init(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[3];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value charname;
  mrb_msagent* agent;
  BSTR name;
  long char_id = 0;

  mrb_get_args(mrb, "|S", &charname);

  charname = mrb_nil_value();
  mrb_get_args(mrb, "|S", &charname);

  agent = malloc(sizeof(mrb_msagent));
  memset(agent, 0, sizeof(mrb_msagent));
  
  hr = CoCreateInstance( &CLSID_AgentServer, NULL, CLSCTX_SERVER,
    &IID_IAgentEx, (LPVOID *)&agent->pAgentEx);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  hr = CoCreateInstance(
    &CLSID_AgentServer,
    NULL,
    CLSCTX_SERVER,
    &IID_IAgentEx,
    (LPVOID *)&agent->pAgentEx);

  // Load
  name = SysAllocString(L"Load");
  hr = agent->pAgentEx->lpVtbl->GetIDsOfNames(agent->pAgentEx, &IID_NULL,
    &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }
  
  VariantInit(&args[2]);
  if (mrb_nil_p(charname)) {
    V_VT(&args[2]) = VT_EMPTY;
  } else {
    V_VT(&args[2]) = VT_BSTR;
    V_BSTR(&args[2]) = utf8_to_bstr(RSTRING_PTR(charname));
  }
  
  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[1]) = &char_id;
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[0]) = &agent->request_id;

  param.rgvarg = &args[0];
  param.cArgs = 3;
  hr = agent->pAgentEx->lpVtbl->Invoke(agent->pAgentEx, dispid, &IID_NULL,
    LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result, NULL, NULL);
  if (V_VT(&args[2]) == VT_BSTR) {
    SysFreeString(V_BSTR(&args[2]));
  }
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }
  
  // GetCharacterEx
  name = SysAllocString(L"GetCharacterEx");
  hr = agent->pAgentEx->lpVtbl->GetIDsOfNames(agent->pAgentEx, &IID_NULL,
    &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }
  
  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_I4;
  V_I4(&args[1]) = char_id;
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_DISPATCH | VT_BYREF;
  V_DISPATCHREF(&args[0]) = &agent->pCharacterEx;
  
  param.cArgs = 2;
  hr = agent->pAgentEx->lpVtbl->Invoke(agent->pAgentEx, dispid, &IID_NULL,
    LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result, NULL, NULL);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  agent->mrb = mrb;
  agent->instance = self;
  mrb_iv_set(mrb, self, mrb_intern(mrb, "context"), mrb_obj_value(
    Data_Wrap_Struct(mrb, mrb->object_class,
    &msagent_type, (void*) agent)));

  return self;
}

static mrb_value
mrb_msagent_show(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[2];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value value_context;
  mrb_msagent* agent = NULL;
  mrb_value with_animation = mrb_true_value();
  BSTR name;

  value_context = mrb_iv_get(mrb, self, mrb_intern(mrb, "context"));
  Data_Get_Struct(mrb, value_context, &msagent_type, agent);

  mrb_get_args(mrb, "|o", &with_animation);

  name = SysAllocString(L"Show");
  hr = agent->pCharacterEx->lpVtbl->GetIDsOfNames(agent->pCharacterEx,
    &IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_BOOL;
  if (mrb_type(with_animation) != MRB_TT_TRUE)
    V_BOOL(&args[1]) = VARIANT_TRUE;
  else
    V_BOOL(&args[1]) = VARIANT_FALSE;
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[0]) = &agent->request_id;
  
  param.rgvarg = &args[0];
  param.cArgs = 2;
  hr = agent->pCharacterEx->lpVtbl->Invoke(agent->pCharacterEx, dispid, &IID_NULL,
    LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result,
    NULL, NULL);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  return mrb_fixnum_value(agent->request_id);
}

static mrb_value
mrb_msagent_hide(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[2];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value value_context;
  mrb_msagent* agent = NULL;
  mrb_value with_animation = mrb_true_value();
  BSTR name;

  value_context = mrb_iv_get(mrb, self, mrb_intern(mrb, "context"));
  Data_Get_Struct(mrb, value_context, &msagent_type, agent);

  mrb_get_args(mrb, "|o", &with_animation);

  name = SysAllocString(L"Hide");
  hr = agent->pCharacterEx->lpVtbl->GetIDsOfNames(agent->pCharacterEx,
    &IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_BOOL;
  if (mrb_type(with_animation) != MRB_TT_TRUE)
    V_BOOL(&args[1]) = VARIANT_TRUE;
  else
    V_BOOL(&args[1]) = VARIANT_FALSE;
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[0]) = &agent->request_id;
  
  param.rgvarg = &args[0];
  param.cArgs = 2;
  hr = agent->pCharacterEx->lpVtbl->Invoke(agent->pCharacterEx, dispid, &IID_NULL,
    LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result,
    NULL, NULL);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  return mrb_fixnum_value(agent->request_id);
}

static mrb_value
mrb_msagent_speak(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[3];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value value_context;
  mrb_msagent* agent = NULL;
  mrb_value text;
  BSTR name;

  value_context = mrb_iv_get(mrb, self, mrb_intern(mrb, "context"));
  Data_Get_Struct(mrb, value_context, &msagent_type, agent);

  mrb_get_args(mrb, "S", &text);

  name = SysAllocString(L"Speak");
  hr = agent->pCharacterEx->lpVtbl->GetIDsOfNames(agent->pCharacterEx,
    &IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  VariantInit(&args[2]);
  V_VT(&args[2]) = VT_BSTR;
  V_BSTR(&args[2]) = utf8_to_bstr(RSTRING_PTR(text));
  
  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_BSTR;
  V_BSTR(&args[1]) = NULL;
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[0]) = &agent->request_id;
  
  param.rgvarg = &args[0];
  param.cArgs = 3;
  hr = agent->pCharacterEx->lpVtbl->Invoke(agent->pCharacterEx, dispid,
    &IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result,
    NULL, NULL);
  SysFreeString(V_BSTR(&args[2]));
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  return mrb_fixnum_value(agent->request_id);
}

static mrb_value
mrb_msagent_play(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[2];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value value_context;
  mrb_msagent* agent = NULL;
  BSTR name;
  mrb_value animation;

  value_context = mrb_iv_get(mrb, self, mrb_intern(mrb, "context"));
  Data_Get_Struct(mrb, value_context, &msagent_type, agent);

  mrb_get_args(mrb, "S", &animation);

  name = SysAllocString(L"Play");
  hr = agent->pCharacterEx->lpVtbl->GetIDsOfNames(agent->pCharacterEx,
    &IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_BSTR;
  V_BSTR(&args[1]) = utf8_to_bstr(RSTRING_PTR(animation));
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4 | VT_BYREF;
  V_I4REF(&args[0]) = &agent->request_id;
  
  param.rgvarg = &args[0];
  param.cArgs = 2;
  hr = agent->pCharacterEx->lpVtbl->Invoke(agent->pCharacterEx, dispid,
    &IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result,
    NULL, NULL);
  SysFreeString(V_BSTR(&args[1]));
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  return mrb_fixnum_value(agent->request_id);
}

static mrb_value
mrb_msagent_move(mrb_state* mrb, mrb_value self) {
  HRESULT hr;
  DISPID dispid = 0;
  VARIANTARG args[2];
  VARIANT result;
  DISPPARAMS param = {0};
  mrb_value value_context;
  mrb_msagent* agent = NULL;
  BSTR name;
  mrb_value x, y;

  value_context = mrb_iv_get(mrb, self, mrb_intern(mrb, "context"));
  Data_Get_Struct(mrb, value_context, &msagent_type, agent);

  mrb_get_args(mrb, "ii", &x, &y);

  name = SysAllocString(L"SetPosition");
  hr = agent->pCharacterEx->lpVtbl->GetIDsOfNames(agent->pCharacterEx,
    &IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  SysFreeString(name);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  VariantInit(&args[1]);
  V_VT(&args[1]) = VT_I4;
  V_I4(&args[1]) = mrb_fixnum(x);
  
  VariantInit(&args[0]);
  V_VT(&args[0]) = VT_I4;
  V_I4(&args[0]) = mrb_fixnum(y);
  
  param.rgvarg = &args[0];
  param.cArgs = 2;
  hr = agent->pCharacterEx->lpVtbl->Invoke(agent->pCharacterEx, dispid,
    &IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &param, &result,
    NULL, NULL);
  if (FAILED(hr)) {
    mrb_raise(mrb, E_RUNTIME_ERROR, RSTRING_PTR(emsg(mrb, hr)));
  }

  return mrb_fixnum_value(agent->request_id);
}

void
mrb_mruby_msagent_gem_init(mrb_state* mrb) {
  struct RClass* _class_msagent;

  CoInitialize(NULL);

  _class_msagent = mrb_define_class(mrb, "MSAgent", mrb->object_class);
  mrb_define_method(mrb, _class_msagent, "initialize", mrb_msagent_init, ARGS_REQ(1));
  mrb_define_method(mrb, _class_msagent, "show", mrb_msagent_show, ARGS_OPT(1));
  mrb_define_method(mrb, _class_msagent, "hide", mrb_msagent_hide, ARGS_OPT(1));
  mrb_define_method(mrb, _class_msagent, "speak", mrb_msagent_speak, ARGS_REQ(1));
  mrb_define_method(mrb, _class_msagent, "play", mrb_msagent_play, ARGS_REQ(1));
  mrb_define_method(mrb, _class_msagent, "move", mrb_msagent_move, ARGS_REQ(2));
}

void
mrb_mruby_msagent_gem_final(mrb_state* mrb) {
  CoUninitialize();
}

/* vim:set et ts=2 sts=2 sw=2 tw=0: */
