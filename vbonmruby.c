#include <mruby.h>
#include <mruby/variable.h>
#include <mruby/string.h>
#include <mruby/compile.h>


// set debug flag
#define DEBUG 0


#if DEBUG
  #include <windows.h>
  #define MAX_DEBUG_LEN 4096
  char DEBUGMSG[MAX_DEBUG_LEN];
#endif


__stdcall int vbonmruby_load_string(
		const char *run_str,
		const char *param_str,
		const char *param_split_str,
		char *result,
		const unsigned int MAX_LEN
	) {
	int res_num;
	mrb_state* st1;
	mrb_value res;
	st1 = mrb_open();

	mrb_gv_set(st1, mrb_intern_cstr(st1, "$params"), mrb_str_new_cstr(st1, param_str));
	mrb_gv_set(st1, mrb_intern_cstr(st1, "$paramss"), mrb_str_new_cstr(st1, param_split_str));
	res = mrb_load_string(st1, run_str);

	switch(mrb_type(res)) {
		case MRB_TT_FIXNUM:
			snprintf(result, MAX_LEN-1, "%d", mrb_fixnum(res));
			res_num = 1;
			break;
		case MRB_TT_FLOAT:
			snprintf(result, MAX_LEN-1, "%f", mrb_float(res));
			res_num = 1;
			break;
		case MRB_TT_SYMBOL:
			snprintf(result, MAX_LEN-1, "%s", mrb_sym2name(st1, mrb_symbol(res)));
			res_num = 1;
			break;
		case MRB_TT_STRING:
			snprintf(result, MAX_LEN-1, "%s", RSTRING_PTR(res));
			res_num = 2;
			break;
		case MRB_TT_UNDEF:
			snprintf(result, MAX_LEN-1, "???  mrb_type: %d\n  UNDEF  Syntax Error?", mrb_type(res));
			res_num = -1;
			break;
		default:
			snprintf(result, MAX_LEN-1, "???  mrb_type: %d", mrb_type(res));
			res_num = 0;
			break;
	}

#if DEBUG
	snprintf(DEBUGMSG, MAX_DEBUG_LEN-1,
			"RUN_STR:\n%s"
				"\n--------\n"
			"PARAM_STR:\n  %s"
				"\n--------\n"
			"PARAM_SPLIT_STR:\n  %s"
				"\n--------\n"
			"MRB_TYPE: %d"
				"\n--------\n"
			"RESULT:\n  %s",
				run_str,
				param_str,
				param_split_str,
				mrb_type(res),
				result
		);
	MessageBox(NULL, DEBUGMSG, "DEBUGMSG", 0);
#endif

	mrb_close(st1);
	return res_num;
}
