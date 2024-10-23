use proc_macro::TokenStream;
use quote::quote;
use syn::{
    punctuated::Punctuated, spanned::Spanned, DataStruct, DeriveInput, Field, FieldsNamed, Ident,
    Token,
};

#[proc_macro_derive(IOExcel, attributes(column))]
pub fn excel_io_derive(input: TokenStream) -> TokenStream {
    let st = syn::parse_macro_input!(input as DeriveInput);
    match do_expand(&st) {
        Ok(tokens) => tokens.into(),
        Err(err) => err.to_compile_error().into(),
    }
}

fn do_expand(st: &DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let struct_ident = &st.ident;
    let excel_struct_ident = get_excel_struct_ident(struct_ident);

    let struct_fields = get_struct_fields(st)?;

    let excel_fields = parse_excel_io_stream(struct_fields)?;

    let from_fields = parse_from_excel_to_record_stream(struct_fields)?;

    let write_token_stream = parse_write_fields_stream(struct_fields)?;

    let token_stream = quote! {
        use calamine::Reader;
        #[derive(Debug, serde::Deserialize)]
        pub struct #excel_struct_ident{
            #(#excel_fields),*
        }

        impl #struct_ident {
            pub fn read_excel(file_path: &str,sheet: &str) -> std::result::Result<std::vec::Vec<#struct_ident>, std::boxed::Box<dyn std::error::Error>> {
                let mut workbook = calamine::open_workbook_auto(file_path)?;

                let range = workbook.worksheet_range(sheet).unwrap();

                let mut result_list = std::vec::Vec::new();

                let deserializer = calamine::RangeDeserializerBuilder::new().from_range(&range)?;
                for result in deserializer {
                    let record: #excel_struct_ident = result?;
                    let record: #struct_ident = record.into();
                    result_list.push(record);
                }
                Ok(result_list)
            }

            pub fn write_excel(file_path: &str,sheet: &str,record_list:&[#struct_ident]) -> std::result::Result<(), std::boxed::Box<dyn std::error::Error>> {
                let mut workbook = rust_xlsxwriter::Workbook::new();
                let worksheet = workbook.add_worksheet();
                worksheet.set_name(sheet)?;
                #write_token_stream
                workbook.save(file_path).unwrap();
                Ok(())
            }
        }

        impl From<#excel_struct_ident> for #struct_ident{
            fn from(record:#excel_struct_ident) -> Self {
                Self{
                    #(#from_fields),*
                }
            }
        }
    };
    Ok(token_stream)
}

fn parse_excel_io_stream(
    struct_fields: &StructFields,
) -> syn::Result<Vec<proc_macro2::TokenStream>> {
    let mut token_stream_list = Vec::new();
    for (row, field) in struct_fields.iter().enumerate() {
        let excel_field = get_io_excel_field(field, row)?;
        let ident = excel_field.ident;
        let ty = excel_field.ty;
        let column = excel_field.column_name.unwrap();
        let token = quote! {
            #[serde(rename = #column)]
            pub #ident: #ty
        };
        token_stream_list.push(token);
    }
    Ok(token_stream_list)
}

fn parse_from_excel_to_record_stream(
    struct_fields: &StructFields,
) -> syn::Result<Vec<proc_macro2::TokenStream>> {
    let mut token_stream_list = Vec::new();
    for (row, field) in struct_fields.iter().enumerate() {
        let excel_field = get_io_excel_field(field, row)?;
        let ident = excel_field.ident;
        let token = quote! {
            #ident:record.#ident
        };
        token_stream_list.push(token);
    }
    Ok(token_stream_list)
}

fn parse_write_fields_stream(
    struct_fields: &StructFields,
) -> syn::Result<proc_macro2::TokenStream> {
    let mut header_token_stream = proc_macro2::TokenStream::new();

    let mut row_token_stream = proc_macro2::TokenStream::new();

    for (row, field) in struct_fields.iter().enumerate() {
        let excel_field = get_io_excel_field(field, row)?;
        let ident = excel_field.ident;
        let header = excel_field.column_name.unwrap();

        let header_token = quote! {
            worksheet.write_string(0, #row as u16, #header)?;
        };

        header_token_stream.extend(header_token);

        let row_token = quote! {
            let val = format!("{}",record.#ident);
            worksheet.write_string(row_idx as u32 + 1, #row as u16, &val)?;
        };
        row_token_stream.extend(row_token);
    }
    let token_stream = quote! {
        #header_token_stream
        for (row_idx, record) in record_list.iter().enumerate() {
            #row_token_stream
        }
    };
    Ok(token_stream)
}

fn get_excel_struct_ident(struct_ident: &Ident) -> syn::Ident {
    let struct_literal = struct_ident.to_string();
    let excel_struct_literal = format!("{}IOExcel", struct_literal);
    syn::Ident::new(&excel_struct_literal, struct_ident.span())
}

type StructFields = Punctuated<Field, Token![,]>;
fn get_struct_fields(st: &DeriveInput) -> syn::Result<&StructFields> {
    if let syn::Data::Struct(DataStruct {
        fields:
            syn::Fields::Named(FieldsNamed {
                named: ref struct_fields,
                ..
            }),
        ..
    }) = &st.data
    {
        Ok(struct_fields)
    } else {
        Err(syn::Error::new_spanned(
            st,
            "IOExcel can only use to struct not enum or other !",
        ))
    }
}

// fn get_inner_type_by_target(field: &Field, target: &str) -> Option<syn::Type> {
//     if let Field {
//         ty:
//             syn::Type::Path(syn::TypePath {
//                 path: syn::Path { ref segments, .. },
//                 ..
//             }),
//         ..
//     } = field
//     {
//         if let Some(syn::PathSegment { ref ident, .. }) = segments.first() {
//             let ident_letial = ident.to_string();
//             if ident_letial == target {
//                 if let syn::PathArguments::AngleBracketed(syn::AngleBracketedGenericArguments {
//                     ref args,
//                     ..
//                 }) = segments.first().unwrap().arguments
//                 {
//                     if let Some(syn::GenericArgument::Type(ty)) = args.first() {
//                         return Some(ty.clone());
//                     }
//                 }
//             }
//         }
//     }
//     None
// }

struct IOExcelField {
    index: usize,
    ident: Ident,
    ty: syn::Type,
    column_name: Option<String>,
}

fn get_io_excel_field(field: &Field, index: usize) -> syn::Result<IOExcelField> {
    let ident = &field.ident;
    let ident_letial = &ident.as_ref().unwrap().to_string();
    let excel_ident = Ident::new(&ident_letial, ident.span());
    let type_ = field.ty.clone();
    let mut column_name = None;
    for attr in &field.attrs {
        if let Ok(syn::Meta::List(list)) = attr.parse_meta() {
            if list.path.is_ident("column") {
                for arg in list.nested {
                    if let syn::NestedMeta::Meta(syn::Meta::NameValue(name_value)) = arg {
                        if name_value.path.is_ident("name") {
                            if let syn::Lit::Str(lit_str) = name_value.lit {
                                column_name = Some(lit_str.value());
                                return Ok(IOExcelField {
                                    index,
                                    ident: excel_ident,
                                    ty: type_,
                                    column_name,
                                });
                            }
                        }
                    }
                }
            }
        }
    }
    Ok(IOExcelField {
        index,
        ident: excel_ident,
        ty: type_,
        column_name,
    })
}
